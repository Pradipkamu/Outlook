"""All Office COM access belongs to this worker's STA thread."""
import datetime as dt
import html
import json
import os
import time
import traceback
from core import utcnow, stamp, parse, UTC

TAGS={'mid':'0x1035001F','refs':'0x1039001F','inreply':'0x1042001F','headers':'0x007D001F'}
OUTLOOK_DISCONNECT_HRESULTS={
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE: Outlook closed/restarted.
    -2147417848,  # RPC_E_DISCONNECTED: COM object disconnected from Outlook.
    -2147023170,  # RPC_S_CALL_FAILED: Outlook terminated during a COM call.
}

def outlook_disconnected(error):
    """Only known transport-loss HRESULTs are safe to reconnect without latching Stop."""
    return getattr(error,'hresult',None) in OUTLOOK_DISCONNECT_HRESULTS

def prop(item,name):
    try:
        p=item.UserProperties.Find(name)
        return str(p.Value) if p else ''
    except Exception:return ''

def snapshot(item,direction='0'):
    if item.Class!=43:return None
    d=dict(entry=str(item.EntryID),store=str(item.Parent.StoreID),conv=str(item.ConversationID or ''),
           subject=str(item.Subject or ''),sent=direction,fid=prop(item,'FOID'),
           occ=prop(item,'FOOccurrence'),generated=prop(item,'FOGenerated'))
    for key,tag in TAGS.items():
        try:d[key]=str(item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/'+tag))
        except Exception:d[key]=''
    d['automatic']='1' if ('auto-submitted: auto-' in d.pop('headers','').lower() or 'automaticreply' in str(item.MessageClass).lower()) else ''
    try:d['received']=stamp(item.ReceivedTime)
    except Exception:d['received']=stamp(utcnow())
    return d

def mark(item,key,value):
    p=item.UserProperties.Find(key)
    if not p:p=item.UserProperties.Add(key,1,False)
    p.Value=str(value)

def heartbeat(store,ready,detail):
    p=store.home/'heartbeat.json';tmp=store.home/'heartbeat.tmp'
    tmp.write_text(json.dumps(dict(time=time.time(),ready=ready,detail=detail)),encoding='utf-8')
    os.replace(tmp,p)

class OutlookWorker:
    def __init__(self,store):
        self.s=store;self.app=None;self.ns=None;self.cursors={};self.scan=None;self.visited=0
        self.pulse_task=None;self.next_pulse=0

    def connect(self):
        import win32com.client
        import pythoncom
        # Do not start hidden Outlook sessions. The user opens Classic Outlook.
        try:active=win32com.client.GetActiveObject('Outlook.Application')
        except pythoncom.com_error as e:
            if e.hresult==-2147221021:return False  # MK_E_UNAVAILABLE: Outlook not open yet.
            raise
        # GetActiveObject already returns a usable dynamic dispatch object.
        # EnsureDispatch depends on generated wrappers under the temporary
        # gen_py cache, which Windows cleanup can remove while Outlook runs.
        self.app=active;self.ns=self.app.GetNamespace('MAPI')
        self.pulse_task=None;self.next_pulse=0
        return True

    def clear_outlook_connection(self):
        """Discard every proxy derived from the old Outlook process."""
        self.scan=None;self.pulse_task=None;self.ns=None;self.app=None

    def pulse_vba(self,clock=None):
        """A non-reminding task change raises VBA ItemChange, never Reminder."""
        now=time.monotonic() if clock is None else clock
        if now<self.next_pulse:return
        self.next_pulse=now+30
        if self.pulse_task is None:
            items=self.ns.GetDefaultFolder(13).Items.Restrict("[Subject] = 'Follow-up Organizer background monitor'")
            if items.Count>50:raise RuntimeError('Background monitor duplicates exceed the safety limit.')
            for i in range(1,items.Count+1):
                candidate=items.Item(i)
                if prop(candidate,'FOBackgroundPulse')=='1':self.pulse_task=candidate;break
            if self.pulse_task is None:
                self.pulse_task=self.app.CreateItem(3)
                self.pulse_task.Subject='Follow-up Organizer background monitor'
                mark(self.pulse_task,'FOBackgroundPulse','1')
        self.pulse_task.ReminderSet=False
        self.pulse_task.Body='Internal background check. No reminder is scheduled. Last check: '+stamp(utcnow())
        self.pulse_task.Save()

    def scan_messages(self):
        """Generator yields after every item/folder. A full pass gates dispatch."""
        since=self.s.meta('scan_since') or stamp(utcnow()-dt.timedelta(days=2))
        cutoff=parse(since)-dt.timedelta(minutes=10)
        roots=[]
        for i in range(1,self.ns.Stores.Count+1):
            st=self.ns.Stores.Item(i)
            for k in (6,5):
                try:roots.append((st.GetDefaultFolder(k),'1' if k==5 else '0'))
                except Exception:pass
        for r in self.s.all():
            if r['status'] in ('Closed','Removed'):continue
            a=r['anchor']
            if a.get('entry'):
                try:roots.append((self.ns.GetItemFromID(a['entry'],a['store']).Parent,a.get('sent','0')))
                except Exception:
                    with self.s.tx():
                        r=self.s.get(r['id'])
                        if r['status']=='Active':
                            r.update(status='Needs action',error='Original mail moved or missing. Link the selected original to refresh its location.',version=r['version']+1)
                            self.s.save(r)
        for f in self.s.meta('folders') or []:
            roots.append((self.ns.GetFolderFromID(f['entry'],f['store']),'0'))
        seen=set();items_seen=0
        for folder,direction in roots:
            key=str(folder.StoreID)+'|'+str(folder.EntryID)
            if key in seen:continue
            seen.add(key)
            if len(seen)>100:raise RuntimeError('Folder scan ceiling reached (100). Reduce monitored folders.')
            # DASL UTC timestamps avoid locale-dependent date interpretation.
            filt='@SQL="http://schemas.microsoft.com/mapi/proptag/0x30080040" >= \''+cutoff.astimezone(UTC).strftime('%Y-%m-%d %H:%M:%S')+'\''
            collection=folder.Items.Restrict(filt)
            count=collection.Count
            if count>10000:raise RuntimeError('Sync backlog exceeds 10,000 items in a folder. Review scan scope.')
            for i in range(count,0,-1):
                if i>collection.Count:
                    yield
                    continue
                item=collection.Item(i)
                m=snapshot(item,direction)
                if m:
                    matched=self.s.ingest(m)
                    if not matched and str(folder.Name)=='Follow-up Mail':
                        with self.s.tx():
                            queue=self.s.meta('setup') or []
                            if not any(x['entry']==m['entry'] and x['store']==m['store'] for x in queue):
                                if len(queue)>=1000:raise RuntimeError('Follow-up setup queue reached 1,000 items. Review pending enrollment.')
                                queue.append(m);self.s.setmeta('setup',queue)
                items_seen+=1
                if items_seen>25000:raise RuntimeError('Scan ceiling reached. Review mailbox scope.')
                yield
            yield

    def process(self,r,n):
        s=self.s
        try:
            with s.tx():
                row=s.db.execute('SELECT status FROM occurrences WHERE fid=? AND n=?',(r['id'],n)).fetchone()
                if not row or row[0]!='Claimed':return
                if s.blocked() or s.get(r['id'])['status']!='Active':
                    raise RuntimeError('Processing stopped before draft creation. Review occurrence before resuming.')
            anchor=self.ns.GetItemFromID(r['anchor']['entry'],r['anchor']['store'])
            mail=anchor.ReplyAll()
            count=mail.Recipients.Count
            for i in range(count,0,-1):mail.Recipients.Remove(i)
            for kind,key in ((1,'to'),(2,'cc')):
                for addr in r.get(key,'').split(';'):
                    if addr:
                        recipient=mail.Recipients.Add(addr);recipient.Type=kind
            if not mail.Recipients.ResolveAll():raise RuntimeError('Could not resolve recipients.')
            account=None
            for i in range(1,self.ns.Accounts.Count+1):
                a=self.ns.Accounts.Item(i)
                if str(a.SmtpAddress).lower()==r['account'].lower():account=a;break
            if account is None:raise RuntimeError('Sending account is unavailable.')
            mail.SendUsingAccount=account
            if str(mail.SendUsingAccount.SmtpAddress).lower()!=r['account'].lower():
                raise RuntimeError('Outlook did not retain the requested sending account.')
            steps=[r['message']]
            if r.get('later_message'):
                steps+=r['later_message'].replace('\r\n','\n').split('\n---STEP---\n')
            message=steps[min(n-1,len(steps)-1)]
            # Prepend through HTMLBody. Writing Body converts Outlook's HTML reply
            # to plain text and can expose signature/CID images as ~WRD*.jpg files.
            original_html=str(mail.HTMLBody or '')
            reminder_html='<div>'+html.escape(message).replace('\r\n','\n').replace('\r','\n').replace('\n','<br>')+'</div><br>'
            mail.HTMLBody=reminder_html+original_html
            mark(mail,'FOID',r['id']);mark(mail,'FOOccurrence',n);mark(mail,'FOGenerated','1')
            mail.Save()
            m=snapshot(mail) or {};m['entry']=str(mail.EntryID);m['store']=str(mail.Parent.StoreID)
            if r['mode']=='Draft':
                with s.tx():
                    current=s.get(r['id'])
                    s.db.execute('UPDATE occurrences SET status=?,data=? WHERE fid=? AND n=?',('Draft',json.dumps(m),r['id'],n))
                    if current['status']=='Active':current['status']='Draft ready';current['version']+=1;s.save(current)
                    s.link(r['id'],m);s.log(r['id'],'Draft ready','Review draft in Outlook before sending.')
                return
            # No DB lock across COM Send: VBA ItemSend must be able to read journal.
            with s.tx():
                if not s.dispatch_allowed(r['id'],n):
                    raise RuntimeError('Dispatch blocked. A draft may exist; review before resuming.')
            if s.blocked():raise RuntimeError('Emergency stop observed before submission.')
            mail.Send()
            # Send return is not reliable sent-copy confirmation. Scanner confirms it.
            with s.tx():s.log(r['id'],'Send requested',f'Occurrence {n}; awaiting Sent Items confirmation')
        except Exception as e:
            s.fail(r['id'],n,str(e))

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        failures=0;next_scan=0
        # Claimed work from a prior process is never retried blindly.
        with self.s.tx():
            for row in self.s.db.execute("SELECT fid,n FROM occurrences WHERE status='Claimed'").fetchall():
                self.s.fail(row[0],row[1],'Restart with unfinished send claim. Check Drafts, Outbox and Sent Items.')
        while True:
            try:
                busy=self.s.meta('ui_busy_until') or 0
                if busy:
                    if busy>time.time():
                        heartbeat(self.s,False,'Waiting for before-send follow-up settings')
                        time.sleep(2);continue
                    self.s.stop('Before-send settings remained open for more than ten minutes. Review before resuming.')
                    self.s.setmeta('ui_busy_until',0)
                if self.app is None and not self.connect():
                    heartbeat(self.s,False,'Waiting for Classic Outlook to open')
                    time.sleep(5);continue
                pythoncom.PumpWaitingMessages()
                # Never hold the DB lock while Save invokes the VBA HTTP callback.
                self.pulse_vba()
                if self.scan is None:
                    self.scan=self.scan_messages();self.pass_start=stamp(utcnow())
                done=False
                for _ in range(50):
                    try:next(self.scan)
                    except StopIteration:done=True;self.scan=None;break
                heartbeat(self.s,False,'Scanning Outlook folders')
                if done:
                    self.s.setmeta('scan_since',self.pass_start)
                    with self.s.tx():
                        for row in self.s.db.execute("SELECT fid,n FROM occurrences WHERE status='Claimed' AND created<?",(stamp(utcnow()-dt.timedelta(minutes=2)),)).fetchall():
                            self.s.fail(row[0],row[1],'Sent copy not confirmed within two minutes. Review Outbox, Drafts and Sent Items.')
                    if not self.app.Session.Offline:
                        self.s.setmeta('last_successful_scan',stamp(utcnow()))
                        heartbeat(self.s,True,'Scan complete; Outlook online')
                        vba=self.s.home/'vba-heartbeat'
                        if vba.exists() and time.time()-vba.stat().st_mtime<180:
                            for r in self.s.all():
                                work=self.s.claim(r['id'])
                                if work:self.process(*work);break
                    failures=0
                time.sleep(1 if not done else 5)
            except Exception as error:
                failures+=1
                detail=traceback.format_exc()
                with (self.s.home/'worker.log').open('a',encoding='utf-8') as f:f.write(stamp(utcnow())+'\n'+detail+'\n')
                self.clear_outlook_connection()
                if outlook_disconnected(error):
                    # Normal lifecycle event: Outlook was closed, updated or restarted.
                    # No dispatch occurs until a fresh COM session completes a full scan.
                    heartbeat(self.s,False,'Waiting for Classic Outlook to reopen')
                    time.sleep(min(30,5*failures))
                    continue
                self.s.stop('Outlook scan/worker error. Check worker.log and Review and resume.')
                heartbeat(self.s,False,'Outlook scan failed - review required')
                time.sleep(min(30,5*failures))
