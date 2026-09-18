"""Follow-up Organizer 2.0 pilot. Pure state/scheduling layer; no Office imports."""
from __future__ import annotations
import calendar
import contextlib
import datetime as dt
import json
import re
import sqlite3
import threading
import time
import uuid
from pathlib import Path

UTC = dt.timezone.utc
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

def utcnow():
    return dt.datetime.now(UTC)

def stamp(d):
    return d.astimezone(UTC).isoformat(timespec='seconds')

def parse(s):
    d = dt.datetime.fromisoformat(str(s))
    return d.replace(tzinfo=IST) if d.tzinfo is None else d

DEFAULTS = dict(weekdays=[0,1,2,3,4,5], start='09:00', end='18:00', holidays=[],
                auto_enabled=False, catchup_enabled=True, catchup_hours=24,
                morning='09:15', evening='17:30')
DEFAULT_IMPORTANT_CATEGORIES=['Quality','Production','Customer','Vendor','Purchase','Finance','Project','Management','Personal','Other']

def slot(d, settings):
    d = d.astimezone(IST)
    start = dt.time.fromisoformat(settings['start'])
    end = dt.time.fromisoformat(settings['end'])
    if start >= end or not settings['weekdays']:
        raise ValueError('Choose working days and an end time later than start time.')
    for _ in range(740):
        if d.weekday() in settings['weekdays'] and d.date().isoformat() not in settings['holidays']:
            if d.time().replace(tzinfo=None) < start:
                return dt.datetime.combine(d.date(), start, IST)
            if d.time().replace(tzinfo=None) < end:
                return d
        d = dt.datetime.combine(d.date()+dt.timedelta(days=1), start, IST)
    raise ValueError('No working slot in the next two years.')

def following(r, now, settings):
    d = parse(r['next_at']).astimezone(IST)
    n = int(r['every']); unit = r['unit']
    if unit == 'hours':
        d += dt.timedelta(hours=n)
    elif unit == 'weeks':
        d += dt.timedelta(weeks=n)
    elif unit == 'months':
        total = d.year*12+d.month-1+n
        year, month = total//12, total%12+1
        day = min(int(r.get('month_day', d.day)), calendar.monthrange(year,month)[1])
        d = d.replace(year=year, month=month, day=day)
    else:
        count=0
        for _ in range(10000):
            d += dt.timedelta(days=1)
            if d.weekday() in settings['weekdays'] and d.date().isoformat() not in settings['holidays']:
                count+=1
            if count == n:
                break
        else:
            raise ValueError('Working-day interval cannot be calculated.')
    # No replay of missed occurrences after a late manual draft submission.
    return slot(max(d, now.astimezone(IST)+dt.timedelta(minutes=1)), settings)

def following_from(r, base, settings):
    """Schedule one full interval after a catch-up send, never replay missed slots."""
    d=base.astimezone(IST)
    n=int(r['every']);unit=r['unit']
    if unit=='hours':
        d+=dt.timedelta(hours=n)
    elif unit=='weeks':
        d+=dt.timedelta(weeks=n)
    elif unit=='months':
        total=d.year*12+d.month-1+n
        year,month=total//12,total%12+1
        day=min(int(r.get('month_day',d.day)),calendar.monthrange(year,month)[1])
        d=d.replace(year=year,month=month,day=day)
    else:
        count=0
        for _ in range(10000):
            d+=dt.timedelta(days=1)
            if d.weekday() in settings['weekdays'] and d.date().isoformat() not in settings['holidays']:
                count+=1
            if count==n:break
        else:raise ValueError('Working-day interval cannot be calculated.')
    return slot(d,settings)

def working_hours(d,hours,settings):
    remaining=hours*3600;d=slot(d,settings)
    for _ in range(740):
        end=dt.datetime.combine(d.date(),dt.time.fromisoformat(settings['end']),IST)
        available=(end-d).total_seconds()
        if remaining<available:return d+dt.timedelta(seconds=remaining)
        remaining-=available
        d=slot(end,settings)
        if remaining==0:return d
    raise ValueError('Review deadline exceeds work-calendar horizon.')

def addresses(s, required=True):
    parts=[p.strip().lower() for p in str(s).split(';') if p.strip()]
    if required and not parts:
        raise ValueError('At least one recipient is required.')
    if len(parts)>30 or any(not re.fullmatch(r'[^\s<>@;,]+@[^\s<>@;,]+\.[^\s<>@;,]+',p) for p in parts):
        raise ValueError('Use SMTP email addresses separated by semicolons (maximum 30).')
    return ';'.join(dict.fromkeys(parts))

class Store:
    def __init__(self, home):
        self.home=Path(home); self.home.mkdir(parents=True,exist_ok=True)
        self.lock=threading.RLock();self.depth=0
        self.db=sqlite3.connect(self.home/'organizer.sqlite3',check_same_thread=False)
        self.db.row_factory=sqlite3.Row
        self.db.executescript('''PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL;
        CREATE TABLE IF NOT EXISTS reminders(id TEXT PRIMARY KEY, data TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS messages(key TEXT PRIMARY KEY, fid TEXT, data TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS occurrences(fid TEXT, n INTEGER, status TEXT, created TEXT, data TEXT,
            PRIMARY KEY(fid,n));
        CREATE TABLE IF NOT EXISTS events(seq INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT, fid TEXT, action TEXT, detail TEXT);
        CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE IF NOT EXISTS excluded_messages(identity TEXT PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS commands(id TEXT PRIMARY KEY, result TEXT);
        CREATE TABLE IF NOT EXISTS important_mails(id TEXT PRIMARY KEY, data TEXT NOT NULL);
        ''')
        settings=self.meta('settings')
        if settings is None:
            self.setmeta('settings',DEFAULTS)
        else:
            # Add new preferences without requiring a database migration or losing user choices.
            merged=dict(DEFAULTS);merged.update(settings)
            if merged!=settings:self.setmeta('settings',merged)
        if self.meta('initialized') is None:
            self.stop('First start: complete the pilot checks, then use Review and resume.')
            self.setmeta('initialized',True)
        if self.meta('important_categories') is None:
            self.setmeta('important_categories',DEFAULT_IMPORTANT_CATEGORIES)

    @contextlib.contextmanager
    def tx(self):
        with self.lock:
            outer=self.depth==0
            self.depth+=1
            try:
                yield
                if outer:self.db.commit()
            except Exception:
                if outer:self.db.rollback()
                raise
            finally:self.depth-=1

    def meta(self,key):
        with self.lock:
            row=self.db.execute('SELECT value FROM meta WHERE key=?',(key,)).fetchone()
            return json.loads(row[0]) if row else None

    def setmeta(self,key,value):
        with self.tx():
            self.db.execute('INSERT OR REPLACE INTO meta VALUES (?,?)',(key,json.dumps(value)))

    def log(self,fid,action,detail=''):
        self.db.execute('INSERT INTO events(at,fid,action,detail) VALUES(?,?,?,?)',
                        (stamp(utcnow()),fid,action,str(detail)[:3000]))

    def save(self,r):
        self.db.execute('INSERT OR REPLACE INTO reminders VALUES(?,?)',(r['id'],json.dumps(r)))

    def all(self):
        with self.lock:
            return [json.loads(x[0]) for x in self.db.execute('SELECT data FROM reminders')]

    def get(self,fid):
        with self.lock:
            row=self.db.execute('SELECT data FROM reminders WHERE id=?',(fid,)).fetchone()
            if not row: raise ValueError('Follow-up not found.')
            return json.loads(row[0])

    def important_all(self):
        with self.lock:
            return [json.loads(x[0]) for x in self.db.execute('SELECT data FROM important_mails')]

    def important_get(self,item_id):
        with self.lock:
            row=self.db.execute('SELECT data FROM important_mails WHERE id=?',(item_id,)).fetchone()
            if not row:raise ValueError('Important-mail bookmark not found.')
            return json.loads(row[0])

    def _important_messages(self,record):
        return record.get('links') or [{key:record.get(key,'') for key in ('entry','store','mid','conv','subject','received','folder')}]

    def _important_values(self,fields):
        category=str(fields.get('category','')).strip()
        if not category or len(category)>40 or any(ord(ch)<32 for ch in category):
            raise ValueError('Category is required and must not exceed 40 characters.')
        priority=str(fields.get('priority','Normal')).strip().title()
        if priority not in ('Normal','High','Critical'):raise ValueError('Priority must be Normal, High or Critical.')
        tags=[]
        for tag in str(fields.get('tags','')).replace(',', ';').split(';'):
            tag=tag.strip()
            if tag and tag.lower() not in [x.lower() for x in tags]:tags.append(tag)
        if len(tags)>20 or any(len(tag)>30 for tag in tags):raise ValueError('Use at most 20 tags, each up to 30 characters.')
        note=str(fields.get('note','')).strip()
        if len(note)>2000:raise ValueError('Private note must not exceed 2,000 characters.')
        review_date=str(fields.get('review_date','')).strip()
        if review_date:dt.date.fromisoformat(review_date)
        categories=list(self.meta('important_categories') or DEFAULT_IMPORTANT_CATEGORIES)
        match=next((x for x in categories if x.lower()==category.lower()),None)
        if match:category=match
        else:
            if len(categories)>=100:raise ValueError('The 100-category limit has been reached.')
            categories.append(category);self.setmeta('important_categories',categories)
        return dict(category=category,priority=priority,tags=';'.join(tags),note=note,review_date=review_date)

    def important_add(self,fields,message,now=None):
        now=now or utcnow()
        with self.tx():
            if not message.get('entry') or not message.get('store'):
                raise ValueError('Save and select an existing Outlook email first.')
            identities=self.message_identities(message)
            for old in self.important_all():
                old_identities=[]
                for linked in self._important_messages(old):old_identities.extend(self.message_identities(linked))
                if set(identities) & set(old_identities):
                    raise ValueError('This Outlook email is already saved in Important Mail.')
            values=self._important_values(fields)
            received=message.get('received') or stamp(now)
            received=stamp(parse(received))
            linked=dict(entry=message['entry'],store=message['store'],mid=message.get('mid',''),
                        conv=message.get('conv',''),subject=str(message.get('subject',''))[:500],
                        received=received,folder=str(message.get('folder',''))[:500])
            record=dict(id=str(uuid.uuid4()),version=1,status='Saved',created=stamp(now),links=[linked],
                        entry=message['entry'],store=message['store'],mid=message.get('mid',''),
                        conv=message.get('conv',''),subject=str(message.get('subject',''))[:500],
                        received=received,folder=str(message.get('folder',''))[:500],**values)
            self.db.execute('INSERT INTO important_mails VALUES(?,?)',(record['id'],json.dumps(record)))
            return record

    def important_attach(self,item_id,message):
        with self.tx():
            record=self.important_get(item_id)
            if not message.get('entry') or not message.get('store'):raise ValueError('Select a saved Outlook email.')
            if message.get('generated'):raise ValueError('Generated reminder messages cannot be linked as Important Mail.')
            identities=set(self.message_identities(message))
            for old in self.important_all():
                for linked in self._important_messages(old):
                    if identities & set(self.message_identities(linked)):
                        return dict(linked=False,duplicate=True,count=len(self._important_messages(record)))
            links=self._important_messages(record)
            if len(links)>=500:raise ValueError('Important Mail link limit reached.')
            received=stamp(parse(message.get('received') or stamp(utcnow())))
            links.append(dict(entry=message['entry'],store=message['store'],mid=message.get('mid',''),
                              conv=message.get('conv',''),subject=str(message.get('subject',''))[:500],
                              received=received,folder=str(message.get('folder',''))[:500]))
            record['links']=links;record['version']+=1
            self.db.execute('UPDATE important_mails SET data=? WHERE id=?',(json.dumps(record),item_id))
            return dict(linked=True,duplicate=False,count=len(links))

    def important_edit(self,item_id,fields,version):
        with self.tx():
            record=self.important_get(item_id)
            if int(version)!=record['version']:raise ValueError('Important-mail bookmark changed. Refresh and try again.')
            record.update(self._important_values(fields));record['version']+=1
            self.db.execute('UPDATE important_mails SET data=? WHERE id=?',(json.dumps(record),item_id))
            return record

    def important_action(self,item_id,action):
        with self.tx():
            record=self.important_get(item_id)
            if action=='Archive':record['status']='Saved' if record['status']=='Archived' else 'Archived';record['version']+=1
            elif action=='Missing':record['status']='Missing';record['version']+=1
            elif action=='Delete':
                self.db.execute('DELETE FROM important_mails WHERE id=?',(item_id,));return
            else:raise ValueError('Unknown Important Mail action.')
            self.db.execute('UPDATE important_mails SET data=? WHERE id=?',(json.dumps(record),item_id))

    def important_relink(self,item_id,message):
        with self.tx():
            record=self.important_get(item_id)
            if not message.get('entry') or not message.get('store'):raise ValueError('Select a saved Outlook email.')
            for old in self.important_all():
                if old['id']!=item_id:
                    for linked in self._important_messages(old):
                        if set(self.message_identities(message)) & set(self.message_identities(linked)):
                            raise ValueError('The selected Outlook email is already bookmarked elsewhere.')
            linked=dict(entry=message['entry'],store=message['store'],mid=message.get('mid',''),
                        conv=message.get('conv',''),subject=str(message.get('subject',''))[:500],
                        received=stamp(parse(message.get('received') or stamp(utcnow()))),
                        folder=str(message.get('folder',''))[:500])
            links=self._important_messages(record)
            if links:links[0]=linked
            else:links=[linked]
            record.update(entry=message['entry'],store=message['store'],mid=message.get('mid',''),
                          conv=message.get('conv',''),subject=str(message.get('subject',''))[:500],
                          received=linked['received'],folder=linked['folder'],links=links,
                          status='Saved',version=record['version']+1)
            self.db.execute('UPDATE important_mails SET data=? WHERE id=?',(json.dumps(record),item_id))
            return record

    def stop(self,reason):
        # Independent file latch also works if DB/worker is blocked.
        (self.home/'STOP').write_text(str(reason),encoding='utf-8')

    def blocked(self):
        p=self.home/'STOP'
        if not p.exists():return ''
        try:return p.read_text(encoding='utf-8') or 'Automation stopped.'
        except OSError:return 'Automation stopped; stop reason unavailable.'

    def health(self):
        try:
            hb=json.loads((self.home/'heartbeat.json').read_text())
            age=time.time()-hb['time']
            return age<45, hb
        except (OSError,ValueError,KeyError): return False,{}

    def resume(self):
        with self.tx():
            good,hb=self.health()
            if not good or not hb.get('ready'): raise ValueError('Outlook scan is not healthy yet. Check worker.log.')
            if not (self.home/'vba-heartbeat').exists() or time.time()-(self.home/'vba-heartbeat').stat().st_mtime>180:
                raise ValueError('VBA monitor is not running. Run FollowupStart in Outlook.')
            (self.home/'STOP').unlink(missing_ok=True)
            self.log('','Resume','User reviewed global safety stop')

    def validate(self,fields,now):
        f=dict(fields)
        f['to']=addresses(f.get('to',''));f['cc']=addresses(f.get('cc',''),False)
        f['account']=addresses(f.get('account',''))
        if ';' in f['account']: raise ValueError('Choose exactly one sending account.')
        if f['account'] in (f['to']+';'+f['cc']).split(';'): raise ValueError('Remove your sending address from recipients.')
        if not f.get('subject','').strip() or not f.get('message','').strip(): raise ValueError('Subject and reminder message are required.')
        if len(f['message'])>20000 or len(f.get('later_message',''))>20000: raise ValueError('Message exceeds 20,000 characters.')
        if f.get('later_message') and any(not x.strip() for x in f['later_message'].replace('\r\n','\n').split('\n---STEP---\n')):
            raise ValueError('Every step in the reminder sequence needs message text.')
        f['every']=int(f.get('every',1));f['maximum']=int(f.get('maximum',5))
        if not 1<=f['every']<=365 or not 1<=f['maximum']<=100: raise ValueError('Interval 1–365; maximum reminders 1–100.')
        if f.get('unit') not in ('hours','working days','weeks','months'): raise ValueError('Invalid interval unit.')
        if f.get('mode') not in ('Draft','Auto'): raise ValueError('Choose Draft or Auto.')
        d=parse(f['next_at'])
        if d<=now: raise ValueError('Choose a future reminder date and time.')
        f['month_day']=d.day;f['next_at']=stamp(slot(d,self.meta('settings')))
        for key in ('deadline','stop_date'):
            if f.get(key): dt.date.fromisoformat(f[key])
        if f.get('stop_date') and parse(f['next_at']).astimezone(IST).date()>dt.date.fromisoformat(f['stop_date']):
            raise ValueError('First working reminder slot falls after the stop date.')
        return f

    def enroll(self,fields,anchor,now=None):
        now=now or utcnow()
        with self.tx():
            f=self.validate(fields,now)
            if anchor.get('pending')!='1' and (not anchor.get('entry') or not anchor.get('store')):
                raise ValueError('Save and select an existing received/sent email before enrollment.')
            for old in self.all():
                if old['status'] not in ('Closed','Removed') and ((anchor.get('entry') and old['anchor'].get('entry')==anchor['entry'] and old['anchor'].get('store')==anchor.get('store')) or
                    (anchor.get('conv') and old['anchor'].get('conv')==anchor['conv'] and old['anchor'].get('store')==anchor.get('store'))):
                    raise ValueError('This conversation already has an open follow-up. Open it in the dashboard.')
            f.update(id=str(uuid.uuid4()),version=1,status='Pending send' if anchor.get('pending')=='1' else 'Active',
                     created=stamp(now),count=0,review_after='',error='',anchor=anchor)
            self.save(f);self.log(f['id'],'Enrolled',f['status'])
            if anchor.get('entry'): self.link(f['id'],anchor)
            queue=self.meta('setup') or []
            self.setmeta('setup',[x for x in queue if not (x.get('entry')==anchor.get('entry') and x.get('store')==anchor.get('store'))])
            return f

    def edit(self,fid,fields,version,now=None):
        with self.tx():
            r=self.get(fid)
            if int(version)!=r['version']: raise ValueError('Record changed. Refresh and edit again.')
            outstanding=self.db.execute("SELECT 1 FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchone()
            if outstanding: raise ValueError('Review the outstanding draft/send outcome first. Use Pause/Close or Resolve occurrence.')
            v=self.validate(fields,now or utcnow())
            for k in ('to','cc','account','subject','message','later_message','next_at','every','unit','maximum','mode','priority','deadline','stop_date','note','month_day'):
                r[k]=v.get(k,'')
            r['status']='Pending send' if r['status']=='Pending send' or not r['anchor'].get('entry') else 'Active'
            r['version']+=1;r['review_after']='';r['error']=''
            self.save(r);self.log(fid,'Schedule saved');return r

    def action(self,fid,action):
        with self.tx():
            r=self.get(fid)
            if action not in ('Pause','Close','Later','Remove'): raise ValueError('Unknown action.')
            if action=='Later':
                if r['status']!='Review':raise ValueError('No reply awaiting review.')
                self.log(fid,'Review later','Reminders remain paused');return
            if action=='Remove':r['status']='Removed'
            elif action=='Close':r['status']='Closed'
            else:r['status']='Paused'
            r['version']+=1
            self.save(r);self.log(fid,action,'Existing drafts/Outbox items need manual review; submitted mail cannot be recalled.')

    def restore_removed(self,fid):
        """Recover a removed follow-up in Paused state; never resume automatically."""
        with self.tx():
            r=self.get(fid)
            if r['status']!='Removed':raise ValueError('Only a removed follow-up can be restored.')
            r.update(status='Paused',review_after='',error='',version=r['version']+1)
            self.save(r);self.log(fid,'Restored','Restored as Paused; review schedule and outstanding drafts before continuing.')
            return r

    def purge_removed(self,fid):
        """Permanently erase one already-removed follow-up and its local audit data."""
        with self.tx():
            r=self.get(fid)
            if r['status']!='Removed':raise ValueError('Only a removed follow-up can be permanently purged.')
            self.db.execute('DELETE FROM occurrences WHERE fid=?',(fid,))
            self.db.execute('DELETE FROM messages WHERE fid=?',(fid,))
            self.db.execute('DELETE FROM events WHERE fid=?',(fid,))
            self.db.execute('DELETE FROM reminders WHERE id=?',(fid,))

    def snooze(self,fid,choice,now=None):
        now=now or utcnow()
        with self.tx():
            r=self.get(fid)
            if r['status'] not in ('Active','Paused'):
                raise ValueError('Snooze requires Active or Paused. Review replies or resolve outstanding sends first.')
            if self.db.execute("SELECT 1 FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchone():
                raise ValueError('Resolve the outstanding draft or send outcome before snoozing.')
            prefs=self.meta('settings');local=now.astimezone(IST)
            if choice=='hour':target=slot(local+dt.timedelta(hours=1),prefs)
            elif choice=='tomorrow':target=slot(local+dt.timedelta(days=1),prefs)
            elif choice=='workday':
                target=slot(dt.datetime.combine(local.date()+dt.timedelta(days=1),dt.time.fromisoformat(prefs['start']),IST),prefs)
            else:raise ValueError('Unknown snooze choice.')
            if target<=parse(r['next_at']):
                raise ValueError('This choice is earlier than the current schedule. Use Edit / Continue to choose another time.')
            if r['count']>=r['maximum'] or (r.get('stop_date') and target.date().isoformat()>r['stop_date']):
                raise ValueError('Snooze would exceed the reminder limit or stop date.')
            r['next_at']=stamp(target);r['version']+=1
            self.save(r);self.log(fid,'Snoozed',r['next_at']+'; deadline and paused state preserved')
            return r

    @staticmethod
    def message_identities(m):
        values=[]
        if m.get('mid'):values.append('mid:'+m['mid'])
        if m.get('entry') and m.get('store'):values.append('loc:'+m['store']+'|'+m['entry'])
        return values

    def excluded(self,m):
        return any(self.db.execute('SELECT 1 FROM excluded_messages WHERE identity=?',(x,)).fetchone() for x in self.message_identities(m))

    def manual_link(self,fid,m):
        with self.tx():
            r=self.get(fid)
            if m.get('generated') or m.get('occ'):raise ValueError('Generated reminder messages cannot change the reply source.')
            if self.db.execute("SELECT 1 FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchone():
                raise ValueError('Resolve the outstanding draft/send outcome before changing links.')
            if not m.get('entry') or not m.get('store'):raise ValueError('Save the Outlook message before linking.')
            self.link(fid,m)
            for identity in self.message_identities(m):self.db.execute('DELETE FROM excluded_messages WHERE identity=?',(identity,))
            r['anchor']=m;r['version']+=1;self.save(r);self.log(fid,'Manual link','Reply source updated; schedule and review state preserved.')

    def attach_related(self,fid,m):
        """Link a user-approved related email without changing the reply source."""
        with self.tx():
            r=self.get(fid)
            if m.get('generated') or m.get('occ'):
                raise ValueError('Generated reminder messages cannot be attached as related Inbox mail.')
            if not m.get('entry') or not m.get('store'):
                raise ValueError('Save the Outlook message before linking it.')
            added=self.link(fid,m)
            if added:
                for identity in self.message_identities(m):
                    self.db.execute('DELETE FROM excluded_messages WHERE identity=?',(identity,))
                r['version']+=1;self.save(r)
                self.log(fid,'Related email linked',m.get('subject',''))
            return dict(r,linked=added)

    def unlink(self,fid,entry,store,version,replacement_entry='',replacement_store=''):
        with self.tx():
            r=self.get(fid)
            if int(version)!=r['version']:raise ValueError('Follow-up changed. Reopen Unlink email and review again.')
            if r['status']=='Pending send':raise ValueError('Wait for the original email to appear in Sent Items before unlinking.')
            if self.db.execute("SELECT 1 FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchone():
                raise ValueError('Resolve the outstanding draft/send outcome before unlinking.')
            found=[(row[0],json.loads(row[1])) for row in self.db.execute('SELECT key,data FROM messages WHERE fid=?',(fid,))]
            targets=[(key,m) for key,m in found if m.get('entry')==entry and m.get('store')==store]
            if len(targets)!=1:raise ValueError('Select one currently linked message.')
            key,m=targets[0]
            if m.get('generated') or m.get('occ'):raise ValueError('Generated reminder history cannot be unlinked.')
            anchor=r['anchor']
            is_anchor=bool(set(self.message_identities(anchor)) & set(self.message_identities(m)))
            if is_anchor:
                candidates=[v for k,v in found if k!=key and v.get('entry')==replacement_entry and v.get('store')==replacement_store and not v.get('generated') and not v.get('occ')]
                if len(candidates)!=1:raise ValueError('This is the reply source. Choose another linked email as its replacement; the only linked email cannot be removed.')
                r['anchor']=candidates[0]
            for identity in self.message_identities(m)+ (self.message_identities(anchor) if is_anchor else []):
                self.db.execute('INSERT OR IGNORE INTO excluded_messages VALUES(?)',(identity,))
            self.db.execute('DELETE FROM messages WHERE key=?',(key,))
            r['version']+=1;self.save(r)
            self.log(fid,'Email unlinked','Original email retained in Outlook; schedule and review state preserved.')
            return r

    def link(self,fid,m):
        key=m.get('mid') or (m.get('store','')+'|'+m.get('entry',''))
        if not key or key=='|': return False
        existing=self.db.execute('SELECT fid FROM messages WHERE key=?',(key,)).fetchone()
        if existing and existing[0]!=fid:raise ValueError('Message belongs to another follow-up.')
        if m.get('generated') and m.get('occ') and m.get('sent')=='1':
            # Replace the obsolete draft locator with its confirmed sent copy.
            for row in self.db.execute('SELECT key,data FROM messages WHERE fid=?',(fid,)).fetchall():
                old=json.loads(row[1])
                if row[0]!=key and old.get('generated') and old.get('occ')==m['occ']:
                    self.db.execute('DELETE FROM messages WHERE key=?',(row[0],))
        self.db.execute('INSERT OR REPLACE INTO messages VALUES(?,?,?)',(key,fid,json.dumps(m)))
        return not bool(existing)

    def ingest(self,m):
        """Idempotent event capture. A reply decision is persisted before notifying VBA."""
        with self.tx():
            if self.excluded(m):return ''
            fid=m.get('fid','');matches=set()
            if fid:
                try:self.get(fid);matches.add(fid)
                except ValueError:pass
            for ref in re.findall(r'<[^<>]+>',m.get('refs','')+' '+m.get('inreply','')):
                row=self.db.execute('SELECT fid FROM messages WHERE key=?',(ref,)).fetchone()
                if row:matches.add(row[0])
            if not matches and m.get('conv'):
                for r in self.all():
                    if r['anchor'].get('conv')==m['conv'] and r['anchor'].get('store')==m.get('store'):matches.add(r['id'])
            if len(matches)!=1:return ''
            fid=matches.pop();r=self.get(fid);fresh=self.link(fid,m)
            if m.get('sent')=='1':
                if r['status']=='Pending send' and m.get('fid')==fid:
                    r['anchor']=m;r['status']='Active';r['version']+=1;self.save(r);self.log(fid,'Original submitted')
                if m.get('occ'):
                    self.submitted(fid,int(m['occ']),m)
                return fid
            if not fresh or m.get('generated') or parse(m.get('received') or stamp(utcnow()))<parse(r['created']):return fid
            if r['status'] in ('Closed','Removed'):
                self.log(fid,'Reply after closure' if r['status']=='Closed' else 'Reply after removal',m.get('subject'));return fid
            # Auto-replies still pause, conservatively; no automatic closure.
            r.update(status='Review',review_after=stamp(working_hours(utcnow(),2,self.meta('settings'))),version=r['version']+1)
            self.save(r);self.log(fid,'Reply received',('Automatic reply: ' if m.get('automatic') else '')+m.get('subject',''))
            return fid

    def submitted(self,fid,n,m):
        r=self.get(fid)
        row=self.db.execute('SELECT status,data FROM occurrences WHERE fid=? AND n=?',(fid,n)).fetchone()
        if not row or row[0]=='Submitted':return
        occurrence=json.loads(row[1] or '{}')
        self.db.execute('UPDATE occurrences SET status=?,data=? WHERE fid=? AND n=?',('Submitted',json.dumps(m),fid,n))
        r['count']=max(r['count'],n)
        r['last_submitted_at']=stamp(utcnow())
        if r['status'] in ('Active','Draft ready'):
            r['status']='Active'
            if occurrence.get('catchup'):
                base=parse(occurrence.get('claimed_at') or stamp(utcnow()))
                r['next_at']=stamp(following_from(r,base,self.meta('settings')))
            else:
                r['next_at']=stamp(following(r,utcnow(),self.meta('settings')))
            if r['count']>=r['maximum']:r['status']='Limit reached'
        r['version']+=1;self.save(r);self.log(fid,'Submitted',f'Occurrence {n}; delivery not confirmed')

    def claim(self,fid,now=None):
        now=now or utcnow()
        with self.tx():
            if self.blocked():return None
            r=self.get(fid)
            due=parse(r['next_at'])
            if r['status']!='Active' or due>now:return None
            settings=self.meta('settings')
            if r['mode']=='Auto' and not settings['auto_enabled']:return None
            if r['count']>=r['maximum'] or (r.get('stop_date') and now.astimezone(IST).date().isoformat()>r['stop_date']):
                r['status']='Limit reached';self.save(r);return None
            overdue=now-due
            catchup=False
            if overdue>dt.timedelta(minutes=10):
                catchup=(r['mode']=='Auto' and settings.get('catchup_enabled',True) and
                         overdue<=dt.timedelta(hours=int(settings.get('catchup_hours',24))))
                if not catchup:
                    if r['mode']=='Auto' and settings.get('catchup_enabled',True):
                        reason=f'Missed schedule exceeds the {int(settings.get("catchup_hours",24))}-hour catch-up window. Edit to choose a future time.'
                    else:
                        reason='Missed schedule. Edit to choose a future time.'
                    r.update(status='Needs action',error=reason,version=r['version']+1)
                    self.save(r);return None
            if slot(now,settings)>now:return None
            n=r['count']+1
            if self.db.execute("SELECT 1 FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchone():return None
            if self.db.execute('SELECT 1 FROM occurrences WHERE fid=? AND n=?',(fid,n)).fetchone():return None
            data=dict(catchup=catchup,scheduled_at=stamp(due),claimed_at=stamp(now))
            self.db.execute('INSERT INTO occurrences VALUES(?,?,?,?,?)',(fid,n,'Claimed',stamp(now),json.dumps(data)))
            self.log(fid,'Catch-up claimed' if catchup else 'Claimed',f'{n}; scheduled {stamp(due)}');return r,n

    def dispatch_allowed(self,fid,n):
        if self.blocked():return False
        r=self.get(fid)
        if r['status']!='Active':return False
        if r['mode']=='Auto' and not self.meta('settings')['auto_enabled']:return False
        now=utcnow()
        for seconds,limit in ((60,5),(3600,50)):
            count=self.db.execute("SELECT COUNT(*) FROM events WHERE action='Dispatch' AND at>=?",(stamp(now-dt.timedelta(seconds=seconds)),)).fetchone()[0]
            if count>=limit:self.stop('Sending ceiling reached. Review before resuming.');return False
        self.log(fid,'Dispatch',str(n));return True

    def fail(self,fid,n,reason):
        with self.tx():
            self.db.execute('UPDATE occurrences SET status=? WHERE fid=? AND n=?',('Uncertain',fid,n))
            r=self.get(fid)
            if r['status'] not in ('Closed','Removed','Review','Paused'):r['status']='Needs action'
            r.update(error=reason,version=r['version']+1);self.save(r)
            self.log(fid,'Send outcome needs review',reason)

    def resolve(self,fid):
        # Explicit user attestation required by UI; consume occurrence, never retry it.
        with self.tx():
            r=self.get(fid)
            rows=self.db.execute("SELECT n FROM occurrences WHERE fid=? AND status IN ('Draft','Claimed','Uncertain')",(fid,)).fetchall()
            if not rows:raise ValueError('No outstanding occurrence.')
            for row in rows:
                self.db.execute('UPDATE occurrences SET status=? WHERE fid=? AND n=?',('Reviewed',fid,row[0]))
                r['count']=max(r['count'],row[0])
            r.update(status='Paused',error='',version=r['version']+1);self.save(r)
            self.log(fid,'User resolved occurrence','User checked Sent Items, Drafts and Outbox. Occurrence consumed; edit to resume.')

    def history(self,fid):
        with self.lock:
            return [dict(r) for r in self.db.execute('SELECT at,action,detail FROM events WHERE fid=? ORDER BY seq DESC LIMIT 100',(fid,))]

    def messages(self,fid):
        with self.lock:
            return [json.loads(r[0]) for r in self.db.execute('SELECT data FROM messages WHERE fid=?',(fid,))]
