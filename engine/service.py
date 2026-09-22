"""Authenticated loopback XML bridge for Outlook VBA. No network mail APIs."""
import argparse
import datetime as dt
import hmac
import json
import os
import secrets
import threading
import time
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from pathlib import Path
from core import Store,stamp,utcnow,DEFAULTS,parse,IST
from paths import data_home,require_migration
from maintenance import maintenance_loop,health_summary,export_diagnostics,backup_snapshot,verify_backup,latest_backup
from version import RELEASE

APPEARANCE={p+'_'+k:v for p in ('dash','edit') for k,v in {'font':'Tahoma','size':'8','color':'Black'}.items()}

HOME=data_home()

def xml_result(values=None,rows=None):
    root=ET.Element('result')
    for key,value in (values or {}).items():ET.SubElement(root,key).text=str(value)
    for r in rows or []:
        row=ET.SubElement(root,'row')
        for key,value in r.items():
            if not isinstance(value,(dict,list)):ET.SubElement(row,key).text=str(value)
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)

def history_rows_ist(rows):
    """Return audit entries newest-first and format timestamps for IST display."""
    result=[]
    for source in sorted(rows,key=lambda row:_date_sort_value(row.get('at')),reverse=True):
        row=dict(source)
        row['at']=parse(row['at']).astimezone(IST).strftime('%d %b %Y %H:%M:%S IST')
        result.append(row)
    return result

def message_rows_ist(rows):
    """Return linked messages newest-first and format received dates in IST."""
    result=[]
    for source in sorted(rows,key=lambda row:_date_sort_value(row.get('received')),reverse=True):
        row=dict(source)
        if row.get('received'):
            row['received']=parse(row['received']).astimezone(IST).strftime('%d %b %Y %H:%M:%S IST')
        result.append(row)
    return result

def _date_sort_value(value):
    """Comparable UTC timestamp; missing or damaged legacy dates sort last."""
    try:return parse(value).astimezone(dt.timezone.utc).timestamp() if value else float('-inf')
    except (TypeError,ValueError,OverflowError):return float('-inf')

class Bridge:
    def __init__(self,s):self.s=s
    def handle(self,action,f):
        s=self.s
        if action=='ping':return xml_result({'blocked':s.blocked(),'health':json.dumps(s.health()[1]),'auto':s.meta('settings')['auto_enabled']})
        if action=='health':
            h=health_summary(s);backup=h.pop('backup')
            return xml_result(dict(h,version=RELEASE,blocked=s.blocked(),backup_file=backup.get('file',''),backup_error=backup.get('error',''),backup_at=backup.get('at','')))
        if action=='diagnostics':return xml_result({'path':str(export_diagnostics(s))})
        if action=='backup_now':
            path=backup_snapshot(s.home,daily=False);return xml_result(dict(verify_backup(path),path=str(path)))
        if action=='backup_check':
            path=latest_backup(s.home);return xml_result(dict(verify_backup(path),path=str(path)))
        if action=='important_categories':
            categories=s.meta('important_categories') or []
            return xml_result({'categories':'|'.join(categories),'count':len(categories),'maximum':100})
        if action=='important_list':
            now=utcnow().astimezone(IST).date();rows=[]
            for source in s.important_all():
                row=dict(source)
                row['received_local']=parse(row['received']).astimezone(IST).strftime('%d %b %Y %H:%M IST')
                row['review_due']='1' if row.get('review_date') and dt.date.fromisoformat(row['review_date'])<=now and row['status']=='Saved' else '0'
                rows.append(row)
            rows.sort(key=lambda row:(row['status']!='Saved',row.get('review_date') or '9999',row.get('priority')!='Critical',row.get('priority')!='High',row['received']),reverse=False)
            return xml_result({'count':len(rows),'review_due':sum(r['review_due']=='1' for r in rows)},rows)
        if action=='important_get':
            record=s.important_get(f['id']);record['link_count']=len(s._important_messages(record))
            return xml_result(record)
        if action=='important_links':
            record=s.important_get(f['id'])
            return xml_result({'count':len(s._important_messages(record))},message_rows_ist(s._important_messages(record)))
        if action=='important_add':
            message={k[2:]:v for k,v in f.items() if k.startswith('m_')}
            return xml_result(s.important_add(f,message))
        if action=='important_edit':return xml_result(s.important_edit(f['id'],f,f['version']))
        if action=='important_attach':
            message={k[2:]:v for k,v in f.items() if k.startswith('a_')}
            return xml_result(s.important_attach(f['id'],message))
        if action in ('important_archive','important_delete','important_missing'):
            method={'important_archive':'Archive','important_delete':'Delete','important_missing':'Missing'}[action]
            ids=f['ids'].splitlines()
            if not ids:raise ValueError('Choose at least one Important Mail bookmark.')
            with s.tx():
                for item_id in ids:s.important_get(item_id)
                for item_id in ids:s.important_action(item_id,method)
            return xml_result({'count':len(ids)})
        if action=='important_relink':
            message={k[2:]:v for k,v in f.items() if k.startswith('m_')}
            return xml_result(s.important_relink(f['id'],message))
        if action=='refresh_location':
            message={k[2:]:v for k,v in f.items() if k.startswith('m_')}
            return xml_result(s.refresh_location(f['id'],message))
        if action=='snooze':return xml_result(s.snooze(f['id'],f['choice']))
        if action=='review_marker':
            if 'marker' in f:s.setmeta('review_marker',f['marker'])
            return xml_result({'marker':s.meta('review_marker') or ''})
        if action=='pulse':
            (s.home/'vba-heartbeat').touch();return xml_result({'ok':1})
        if action=='ui_busy':
            s.setmeta('ui_busy_until',time.time()+600 if f.get('busy')=='1' else 0)
            return xml_result({'ok':1})
        if action=='stop':s.stop('User pressed Stop Automation.');return xml_result({'ok':1})
        if action=='resume':s.resume();return xml_result({'ok':1})
        if action=='list':
            rows=[];now=utcnow();today=now.astimezone(IST).date()
            for r in s.all():
                if r['status']=='Removed' and f.get('include_removed')!='1':continue
                d=dict(r);d['anchor_entry']=r['anchor'].get('entry','');d['anchor_store']=r['anchor'].get('store','')
                due=parse(r['next_at']);due_local=due.astimezone(IST)
                d['next_local']=due_local.strftime('%d %b %Y %H:%M IST')
                with s.lock:
                    last=s.db.execute("SELECT at FROM events WHERE fid=? AND action='Submitted' ORDER BY seq DESC LIMIT 1",(r['id'],)).fetchone()
                    count=s.db.execute("SELECT COUNT(*) FROM occurrences WHERE fid=? AND status='Submitted'",(r['id'],)).fetchone()[0]
                    outcome=s.db.execute('SELECT status FROM occurrences WHERE fid=? ORDER BY n DESC LIMIT 1',(r['id'],)).fetchone()
                d['last_local']=parse(last[0]).astimezone(IST).strftime('%d %b %H:%M') if last else '-'
                d['submitted_count']=str(count)
                d['send_state']={'Submitted':'In Sent Items','Claimed':'Awaiting sent copy','Draft':'Draft only','Uncertain':'Check send outcome','Reviewed':'Reviewed'}.get(outcome[0] if outcome else '', 'Not submitted')
                d['next_local']=parse(r['next_at']).astimezone(IST).strftime('%d %b %H:%M') if r['status']=='Active' else '-'
                d['overdue']='1' if r['status']=='Active' and due<now else '0'
                d['due_today']='1' if r['status']=='Active' and due_local.date()==today else '0'
                d['exception']='1' if (r['status']=='Needs action' or bool(r.get('error')) or d['send_state'] in ('Check send outcome','Awaiting sent copy')) else '0'
                rows.append(d)
            for i,m in enumerate(s.meta('setup') or []):
                rows.append(dict(id='setup:'+m['store']+'|'+m['entry'],subject=m.get('subject',''),to='',status='Setup required',next_at='Configure recipients and schedule',anchor_entry=m['entry'],anchor_store=m['store'],overdue='0',due_today='0',exception='1'))
            rank={'Review':0,'Needs action':1,'Setup required':2,'Limit reached':3,'Draft ready':4,'Pending send':5,'Removed':9}
            rows.sort(key=lambda r:(rank.get(r['status'],6),r.get('deadline') or '9999',r.get('priority')!='High',r.get('next_at','')))
            counts=dict(
                overdue=sum(r.get('overdue')=='1' for r in rows),
                due_today=sum(r.get('due_today')=='1' for r in rows),
                awaiting=sum(r.get('status') in ('Active','Draft ready','Paused','Pending send') for r in rows),
                review=sum(r.get('status')=='Review' for r in rows),
                exceptions=sum(r.get('exception')=='1' for r in rows),
                removed=sum(r.get('status')=='Removed' for r in rows))
            return xml_result(dict(blocked=s.blocked(),engine=health_summary(s)['state'],version=RELEASE,**counts),rows)
        if action=='find':
            matches=[]
            for r in s.all():
                if r['status'] in ('Closed','Removed'):continue
                if f.get('fid')==r['id'] or (f.get('conv') and f.get('conv')==r['anchor'].get('conv') and f.get('store')==r['anchor'].get('store')):matches.append(r['id'])
            return xml_result({'id':matches[0] if len(matches)==1 else ''})
        if action=='history':return xml_result(rows=history_rows_ist(s.history(f['id'])))
        if action=='messages':return xml_result(rows=message_rows_ist(s.messages(f['id'])))
        if action=='get':
            if f['id'].startswith('setup:'):
                for m in s.meta('setup') or []:
                    if f['id']=='setup:'+m['store']+'|'+m['entry']:
                        return xml_result(dict(type='setup',anchor_entry=m['entry'],anchor_store=m['store']))
                raise ValueError('Setup item no longer exists. Refresh the dashboard.')
            return xml_result(s.get(f['id']))
        if action=='enroll':
            anchor={k[2:]:v for k,v in f.items() if k.startswith('a_')}
            return xml_result(s.enroll(f,anchor))
        if action=='edit':return xml_result(s.edit(f['id'],f,f['version']))
        if action in ('Pause','Close','Later','Remove'):
            if f['id'].startswith('setup:'):raise ValueError('Configure this email with Edit / Continue before managing its schedule.')
            s.action(f['id'],action);return xml_result({'ok':1})
        if action=='bulk':
            ids=f['ids'].splitlines();method=f['method']
            if method not in ('Pause','Close','Later','Remove') or not ids:raise ValueError('Choose follow-ups and an action.')
            with s.tx():
                for fid in ids:s.get(fid)
                for fid in ids:s.action(fid,method)
            return xml_result({'count':len(ids)})
        if action=='restore':
            ids=f['ids'].splitlines()
            if not ids:raise ValueError('Choose removed follow-ups to restore.')
            with s.tx():
                for fid in ids:
                    if s.get(fid)['status']!='Removed':raise ValueError('Restore selection contains an item that is not Removed.')
                for fid in ids:s.restore_removed(fid)
            return xml_result({'count':len(ids)})
        if action=='purge':
            ids=f['ids'].splitlines()
            if not ids or f.get('confirm')!='PURGE':raise ValueError('Permanent purge confirmation is missing.')
            with s.tx():
                for fid in ids:
                    if s.get(fid)['status']!='Removed':raise ValueError('Purge selection contains an item that is not Removed.')
                for fid in ids:s.purge_removed(fid)
            return xml_result({'count':len(ids)})
        if action=='setup':
            with s.tx():
                if not s.ingest(f):
                    queue=s.meta('setup') or []
                    if not any(x['entry']==f['entry'] and x['store']==f['store'] for x in queue):
                        queue.append(f);s.setmeta('setup',queue)
            return xml_result({'ok':1})
        if action=='resolve':s.resolve(f['id']);return xml_result({'ok':1})
        if action=='event':return xml_result({'id':s.ingest(f)})
        if action=='link':
            s.manual_link(f['id'],{k[2:]:v for k,v in f.items() if k.startswith('a_')})
            return xml_result({'ok':1})
        if action=='attach_related':
            return xml_result(s.attach_related(f['id'],{k[2:]:v for k,v in f.items() if k.startswith('a_')}))
        if action=='unlink':
            s.unlink(f['id'],f['entry'],f['store'],f['version'],f.get('replacement_entry',''),f.get('replacement_store',''))
            return xml_result({'ok':1})
        if action=='appearance':
            return xml_result(s.meta('appearance') or APPEARANCE)
        if action=='save_appearance':
            values={key:f.get(key,default) for key,default in APPEARANCE.items()}
            for panel in ('dash','edit'):
                if values[panel+'_font'] not in ('Tahoma','Calibri','Arial','Segoe UI'):raise ValueError('Choose a listed font.')
                if values[panel+'_size'] not in ('8','9','10'):raise ValueError('Choose font size 8, 9 or 10.')
                if values[panel+'_color'] not in ('Black','Navy','Dark green','Maroon','Purple'):raise ValueError('Choose a listed text color.')
            s.setmeta('appearance',values)
            return xml_result(values)
        if action=='folder':
            folders=s.meta('folders') or [];item={'entry':f['entry'],'store':f['store']}
            if item not in folders:folders.append(item)
            s.setmeta('folders',folders);return xml_result({'ok':1})
        if action=='settings':return xml_result({k:json.dumps(v) if isinstance(v,list) else v for k,v in s.meta('settings').items()})
        if action=='import':
            if not s.blocked():raise ValueError('Press Stop Automation before importing Excel changes.')
            node=ET.fromstring(f['batch']);seen=set();prepared=[]
            with s.tx():
                for row in node:
                    values={c.tag:c.text or '' for c in row};fid=values['id']
                    if fid in seen:raise ValueError('Duplicate ID in workbook.')
                    seen.add(fid);r=s.get(fid)
                    if int(values['version'])!=r['version']:raise ValueError('Workbook is stale. Export a fresh copy.')
                    # Validate changed fields only; an untouched overdue row is legal.
                    keys=('to','cc','account','subject','message','later_message','next_at','every','unit','maximum','mode','priority','deadline','stop_date','note')
                    changed=any(str(values.get(k,''))!=str(r.get(k,'')) for k in keys)
                    if changed:
                        if s.db.execute("SELECT 1 FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchone():
                            raise ValueError('Resolve outstanding draft/send outcome before editing configuration.')
                        valid=s.validate(values,utcnow())
                        for k in keys+('month_day',):r[k]=valid.get(k,'')
                    requested=values.get('status',r['status'])
                    if requested in ('Closed','Paused'):r['status']=requested
                    elif requested!=r['status']:raise ValueError('Excel cannot resume or change review status. Use Outlook Edit / Continue.')
                    r['version']+=1;prepared.append(r)
                for r in prepared:s.save(r);s.log(r['id'],'Excel import','Configuration updated; global stop retained')
            return xml_result({'count':len(prepared)})
        if action=='save_settings':
            settings=dict(s.meta('settings'))
            settings['weekdays']=[int(x) for x in f['weekdays'].split(',')]
            if not settings['weekdays'] or any(x not in range(7) for x in settings['weekdays']):raise ValueError('Working weekdays must be numbers 0 through 6.')
            for key in ('start','end','morning','evening'):dt.time.fromisoformat(f[key]);settings[key]=f[key]
            if settings['start']>=settings['end']:raise ValueError('Working start must precede end.')
            settings['holidays']=[x.strip() for x in f.get('holidays','').split(';') if x.strip()]
            for x in settings['holidays']:dt.date.fromisoformat(x)
            settings['auto_enabled']=f.get('auto_enabled')=='1'
            settings['catchup_enabled']=f.get('catchup_enabled','1' if settings.get('catchup_enabled',True) else '0')=='1'
            catchup_hours=int(f.get('catchup_hours',settings.get('catchup_hours',24)))
            if not 1<=catchup_hours<=168:raise ValueError('Catch-up window must be from 1 to 168 hours.')
            settings['catchup_hours']=catchup_hours
            s.setmeta('settings',settings);return xml_result({'ok':1})
        if action=='permit':
            # Manual sending of generated drafts still respects Close/Review/Stop.
            r=s.get(f['id']);occ=int(f['occ'])
            with s.lock:
                row=s.db.execute('SELECT status FROM occurrences WHERE fid=? AND n=?',(r['id'],occ)).fetchone()
            allowed=(not s.blocked() and r['status'] in ('Active','Draft ready') and row is not None and row[0] in ('Draft','Claimed'))
            return xml_result({'allowed':'1' if allowed else '0'})
        raise ValueError('Unknown command.')

def main():
    require_migration()
    HOME.mkdir(parents=True,exist_ok=True)
    if dt.datetime.now().astimezone().utcoffset()!=dt.timedelta(hours=5,minutes=30):
        raise SystemExit('This pilot requires Windows time zone India Standard Time (UTC+05:30).')
    # OS-owned lock, released on process termination (no stale lockfile bypass).
    import msvcrt
    lockfile=open(HOME/'service.lock','a+b');lockfile.seek(0);lockfile.write(b'1');lockfile.flush();lockfile.seek(0)
    try:msvcrt.locking(lockfile.fileno(),msvcrt.LK_NBLCK,1)
    except OSError:raise SystemExit('Organizer is already running.')
    tokenfile=HOME/'token.txt'
    if not tokenfile.exists():tokenfile.write_text(secrets.token_hex(32),encoding='ascii')
    token=tokenfile.read_text().strip();store=Store(HOME);bridge=Bridge(store)
    class Handler(BaseHTTPRequestHandler):
        def log_message(self,*args):pass
        def do_POST(self):
            if not hmac.compare_digest(self.headers.get('X-FO-Token',''),token):self.send_error(403);return
            if self.headers.get('Origin'):self.send_error(403);return
            try:
                length=int(self.headers.get('Content-Length','0'))
                if not 0<length<=2097152:raise ValueError('Request exceeds 2 MB. Import fewer rows.')
                self.connection.settimeout(5)
                body=self.rfile.read(length)
                if b'<!DOCTYPE' in body.upper() or b'<!ENTITY' in body.upper():raise ValueError('Entities are not accepted.')
                node=ET.fromstring(body);f={c.tag:c.text or '' for c in node};action=self.path.lstrip('/')
                # Only mutations use idempotency tokens; persisted result makes retries safe.
                cid=f.pop('command_id','')
                if cid:
                    with store.tx():
                        prior=store.db.execute('SELECT result FROM commands WHERE id=?',(cid,)).fetchone()
                        if prior:result=prior[0].encode()
                        else:
                            result=bridge.handle(action,f)
                            store.db.execute('INSERT INTO commands VALUES(?,?)',(cid,result.decode()))
                else:result=bridge.handle(action,f)
                self.send_response(200)
            except Exception as e:
                result=xml_result({'error':str(e)});self.send_response(400)
            self.send_header('Content-Type','application/xml; charset=utf-8');self.send_header('Content-Length',str(len(result)))
            self.end_headers();self.wfile.write(result)
    from outlook_worker import OutlookWorker
    threading.Thread(target=OutlookWorker(store).run,daemon=True).start()
    threading.Thread(target=maintenance_loop,args=(store,),daemon=True).start()
    ThreadingHTTPServer(('127.0.0.1',8765),Handler).serve_forever()

if __name__=='__main__':main()
