import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from core import Store,IST,stamp
from outlook_worker import OutlookWorker,outlook_disconnected

class Recipients:
    def __init__(self):self.values=[NS(Type=1)]
    @property
    def Count(self):return len(self.values)
    def Remove(self,i):self.values.pop(i-1)
    def Add(self,address):
        r=NS(address=address,Type=1);self.values.append(r);return r
    def ResolveAll(self):return True

class Properties:
    def __init__(self):self.data={}
    def Find(self,key):return self.data.get(key)
    def Add(self,key,*args):self.data[key]=NS(Value='');return self.data[key]

class Mail:
    def __init__(self):
        self.Recipients=Recipients();self.UserProperties=Properties();self.Body='Original quoted body'
        self.HTMLBody='<html><body>Original <img src="cid:signature-logo"></body></html>'
        self.EntryID='draft';self.Parent=NS(StoreID='store');self.sends=0;self.fail_send=False
    def Save(self):pass
    def Send(self):
        self.sends+=1
        if self.fail_send:raise RuntimeError('COM disconnected after submission attempt')

class WorkerTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.s=Store(self.tmp.name);self.s.home.joinpath('STOP').unlink()
        self.now=dt.datetime(2026,9,14,9,0,tzinfo=IST)
        self.r=self.s.enroll(dict(subject='Readiness',to='supplier@example.com',cc='',account='me@example.com',next_at=stamp(self.now+dt.timedelta(minutes=1)),every=1,unit='working days',maximum=5,mode='Draft',message='First message',later_message='Second message\n---STEP---\nThird message'),dict(entry='original',store='store',conv='thread'),self.now)
        self.mail=Mail();self.worker=OutlookWorker(self.s)
        accounts=NS(Count=1,Item=lambda i:NS(SmtpAddress='me@example.com'))
        self.worker.ns=NS(GetItemFromID=lambda *args:NS(ReplyAll=lambda:self.mail),Accounts=accounts)
        self.patch=patch('outlook_worker.snapshot',lambda item:dict(entry='draft',store='store',generated='1',occ='1',sent='0'));self.patch.start()
    def tearDown(self):self.patch.stop();self.s.db.close();self.tmp.cleanup()
    def claim(self):return self.s.claim(self.r['id'],self.now+dt.timedelta(minutes=2))
    def test_known_outlook_disconnects_are_reconnectable(self):
        for hresult in (-2147023174,-2147417848,-2147023170):
            self.assertTrue(outlook_disconnected(NS(hresult=hresult)))

    def test_unknown_com_errors_still_require_safety_review(self):
        self.assertFalse(outlook_disconnected(NS(hresult=-2147352567)))
        self.assertFalse(outlook_disconnected(RuntimeError('unexpected worker failure')))

    def test_draft_does_not_send_or_advance_count(self):
        self.worker.process(*self.claim())
        r=self.s.get(self.r['id']);self.assertEqual(r['status'],'Draft ready');self.assertEqual(r['count'],0);self.assertEqual(self.mail.sends,0)
        self.assertEqual(self.mail.Recipients.values[0].address,'supplier@example.com')
    def test_stop_between_claim_and_dispatch_blocks_auto(self):
        self.r['mode']='Auto'
        with self.s.tx():self.s.save(self.r)
        prefs=self.s.meta('settings');prefs['auto_enabled']=True;self.s.setmeta('settings',prefs)
        claim=self.claim();self.s.stop('Stop pressed')
        self.worker.process(*claim)
        self.assertEqual(self.mail.sends,0);self.assertEqual(self.s.get(self.r['id'])['status'],'Needs action')
    def test_com_error_does_not_retry_send(self):
        self.r['mode']='Auto'
        with self.s.tx():self.s.save(self.r)
        prefs=self.s.meta('settings');prefs['auto_enabled']=True;self.s.setmeta('settings',prefs)
        claim=self.claim();self.mail.fail_send=True;self.worker.process(*claim)
        self.assertEqual(self.mail.sends,1);self.assertIsNone(self.claim())
    def test_sequence_uses_last_message_after_last_step(self):
        self.r['count']=3
        with self.s.tx():self.s.save(self.r)
        self.worker.process(*self.claim())
        self.assertTrue(self.mail.HTMLBody.startswith('<div>Third message</div><br>'))
        self.assertIn('cid:signature-logo',self.mail.HTMLBody)
        self.assertEqual(self.mail.Body,'Original quoted body')

    def test_html_message_is_escaped_and_line_breaks_are_preserved(self):
        self.r['message']='Status <urgent> & required\r\nPlease reply'
        with self.s.tx():self.s.save(self.r)
        self.worker.process(*self.claim())
        self.assertTrue(self.mail.HTMLBody.startswith('<div>Status &lt;urgent&gt; &amp; required<br>Please reply</div><br>'))
        self.assertIn('cid:signature-logo',self.mail.HTMLBody)
    def test_background_pulse_has_no_reminder_and_is_rate_limited(self):
        saves=[];creates=[]
        task=NS(UserProperties=Properties(),ReminderSet=True,Save=lambda:saves.append(1))
        items=NS(Count=0)
        self.worker.ns=NS(GetDefaultFolder=lambda kind:NS(Items=NS(Restrict=lambda query:items)))
        def create(kind):creates.append(kind);return task
        self.worker.app=NS(CreateItem=create)
        self.worker.pulse_vba(100);self.worker.pulse_vba(101);self.worker.pulse_vba(130)
        self.assertEqual(creates,[3]);self.assertEqual(len(saves),2)
        self.assertFalse(task.ReminderSet)
        self.assertEqual(task.UserProperties.Find('FOBackgroundPulse').Value,'1')
    def test_background_pulse_reuses_only_tagged_task(self):
        saves=[]
        unrelated=NS(UserProperties=Properties())
        task=NS(UserProperties=Properties(),ReminderSet=True,Save=lambda:saves.append(1))
        task.UserProperties.Add('FOBackgroundPulse').Value='1'
        items=NS(Count=2,Item=lambda i:[unrelated,task][i-1])
        self.worker.ns=NS(GetDefaultFolder=lambda kind:NS(Items=NS(Restrict=lambda query:items)))
        self.worker.app=NS(CreateItem=lambda kind:self.fail('Should reuse tagged task'))
        self.worker.pulse_vba(100)
        self.assertEqual(len(saves),1);self.assertFalse(task.ReminderSet)

if __name__=='__main__':unittest.main()
