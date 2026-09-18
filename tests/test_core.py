import datetime as dt
import json
import sys
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from core import Store,IST,UTC,DEFAULTS,slot,following,following_from,working_hours,parse,stamp,utcnow
from service import Bridge

class OrganizerTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.s=Store(self.tmp.name)
        self.now=dt.datetime(2026,9,14,9,0,tzinfo=IST)
        self.fields=dict(subject='Die readiness',to='toolroom@example.com',cc='',account='me@example.com',
                         next_at=stamp(self.now+dt.timedelta(minutes=5)),every='1',unit='working days',maximum='5',
                         mode='Draft',message='Please confirm readiness.',later_message='',priority='High',deadline='',stop_date='',note='')
        self.anchor=dict(entry='original',store='store1',conv='thread1',mid='<original@example.com>',sent='1')
    def tearDown(self):
        self.s.db.close();self.tmp.cleanup()
    def enroll(self):return self.s.enroll(self.fields,self.anchor,self.now)
    def test_work_window_holiday(self):
        prefs=dict(DEFAULTS,weekdays=[0,1,2,3,4],holidays=['2026-09-14'])
        d=dt.datetime(2026,9,11,19,tzinfo=IST)
        self.assertEqual(slot(d,prefs),dt.datetime(2026,9,15,9,tzinfo=IST))
    def test_month_anchor_survives_february(self):
        r=dict(next_at='2027-01-31T10:00:00+05:30',every=1,unit='months',month_day=31)
        prefs=dict(DEFAULTS,weekdays=list(range(7)))
        feb=following(r,dt.datetime(2027,1,31,10,tzinfo=IST),prefs)
        self.assertEqual(feb.day,28)
        r['next_at']=stamp(feb)
        self.assertEqual(following(r,feb,prefs).day,31)
    def test_invalid_recipient_blocks(self):
        self.fields['to']='Tool Room'
        with self.assertRaises(ValueError):self.enroll()
    def test_self_recipient_blocks(self):
        self.fields['to']='me@example.com'
        with self.assertRaises(ValueError):self.enroll()
    def test_duplicate_conversation_blocks(self):
        self.enroll();self.anchor['entry']='another'
        with self.assertRaises(ValueError):self.enroll()
    def test_identical_subject_different_thread_is_separate(self):
        self.enroll();self.anchor.update(entry='another',conv='thread2',mid='<other@example.com>')
        self.enroll();self.assertEqual(len(self.s.all()),2)
    def test_reply_pauses_and_replay_does_not_increment_version(self):
        r=self.enroll()
        m=dict(entry='reply',store='store1',mid='<reply@example.com>',inreply='<original@example.com>',subject='Changed subject',received=stamp(self.now+dt.timedelta(minutes=10)),sent='0')
        self.assertEqual(self.s.ingest(m),r['id'])
        current=self.s.get(r['id']);self.assertEqual(current['status'],'Review')
        self.s.ingest(m);self.assertEqual(self.s.get(r['id'])['version'],current['version'])
        self.s.action(r['id'],'Later');self.assertEqual(self.s.get(r['id'])['status'],'Review')
    def test_closed_reply_does_not_reopen(self):
        r=self.enroll();self.s.action(r['id'],'Close')
        self.s.ingest(dict(entry='reply',store='store1',conv='thread1',sent='0',received=stamp(self.now+dt.timedelta(minutes=10))))
        self.assertEqual(self.s.get(r['id'])['status'],'Closed')
    def test_pending_draft_does_not_schedule(self):
        self.anchor['pending']='1';r=self.enroll();(self.s.home/'STOP').unlink()
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(minutes=6)))
        self.s.ingest(dict(self.anchor,pending='',fid=r['id'],sent='1'))
        self.assertEqual(self.s.get(r['id'])['status'],'Active')
    def test_single_claim_and_reply_prevents_dispatch(self):
        r=self.enroll();(self.s.home/'STOP').unlink()
        now=self.now+dt.timedelta(minutes=6)
        self.assertIsNotNone(self.s.claim(r['id'],now));self.assertIsNone(self.s.claim(r['id'],now))
        self.s.action(r['id'],'Pause')
        self.assertFalse(self.s.dispatch_allowed(r['id'],1))
    def test_stop_survives_new_store(self):
        r=self.enroll();self.s.stop('test stop')
        self.s.db.close();self.s=Store(self.tmp.name)
        self.assertEqual(self.s.blocked(),'test stop');self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(minutes=6)))
    def test_missed_draft_schedule_requires_review(self):
        r=self.enroll();(self.s.home/'STOP').unlink()
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(hours=1)))
        self.assertEqual(self.s.get(r['id'])['status'],'Needs action')
    def test_auto_missed_schedule_claims_one_catchup_within_window(self):
        self.fields['mode']='Auto';r=self.enroll();(self.s.home/'STOP').unlink()
        prefs=self.s.meta('settings');prefs['auto_enabled']=True;self.s.setmeta('settings',prefs)
        now=self.now+dt.timedelta(hours=1)
        work=self.s.claim(r['id'],now)
        self.assertIsNotNone(work);self.assertEqual(work[1],1)
        self.assertIsNone(self.s.claim(r['id'],now))
        row=self.s.db.execute('SELECT data FROM occurrences WHERE fid=? AND n=1',(r['id'],)).fetchone()
        self.assertTrue(json.loads(row[0])['catchup'])
        self.assertEqual(self.s.history(r['id'])[0]['action'],'Catch-up claimed')
    def test_catchup_schedules_one_full_interval_after_actual_claim(self):
        self.fields['mode']='Auto';r=self.enroll();(self.s.home/'STOP').unlink()
        prefs=self.s.meta('settings');prefs['auto_enabled']=True;self.s.setmeta('settings',prefs)
        claimed=self.now+dt.timedelta(hours=1)
        self.s.claim(r['id'],claimed);self.s.submitted(r['id'],1,dict(entry='sent',store='store1'))
        current=self.s.get(r['id'])
        self.assertEqual(parse(current['next_at']).astimezone(IST),dt.datetime(2026,9,15,10,0,tzinfo=IST))
        self.assertEqual(current['count'],1)
        self.assertEqual(self.s.db.execute('SELECT COUNT(*) FROM occurrences WHERE fid=?',(r['id'],)).fetchone()[0],1)
    def test_auto_missed_schedule_beyond_window_requires_review(self):
        self.fields['mode']='Auto';r=self.enroll();(self.s.home/'STOP').unlink()
        prefs=self.s.meta('settings');prefs['auto_enabled']=True;self.s.setmeta('settings',prefs)
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(hours=30)))
        current=self.s.get(r['id'])
        self.assertEqual(current['status'],'Needs action');self.assertIn('24-hour',current['error'])
    def test_disabled_catchup_requires_review(self):
        self.fields['mode']='Auto';r=self.enroll();(self.s.home/'STOP').unlink()
        prefs=self.s.meta('settings');prefs.update(auto_enabled=True,catchup_enabled=False);self.s.setmeta('settings',prefs)
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(hours=1)))
        self.assertEqual(self.s.get(r['id'])['status'],'Needs action')
    def test_reply_scanned_before_catchup_blocks_send(self):
        self.fields['mode']='Auto';r=self.enroll();(self.s.home/'STOP').unlink()
        prefs=self.s.meta('settings');prefs['auto_enabled']=True;self.s.setmeta('settings',prefs)
        reply=dict(entry='reply-before-reopen',store='store1',mid='<reply-before-reopen@example.com>',
                   inreply='<original@example.com>',subject='Reply received while Outlook closed',
                   received=stamp(self.now+dt.timedelta(minutes=30)),sent='0')
        self.s.ingest(reply)
        self.assertEqual(self.s.get(r['id'])['status'],'Review')
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(hours=1)))
        self.assertEqual(self.s.db.execute('SELECT COUNT(*) FROM occurrences WHERE fid=?',(r['id'],)).fetchone()[0],0)
    def test_old_settings_gain_catchup_defaults(self):
        old=dict(DEFAULTS);old.pop('catchup_enabled');old.pop('catchup_hours')
        self.s.setmeta('settings',old);self.s.db.close();self.s=Store(self.tmp.name)
        prefs=self.s.meta('settings')
        self.assertTrue(prefs['catchup_enabled']);self.assertEqual(prefs['catchup_hours'],24)
    def test_catchup_window_validation(self):
        with self.assertRaises(ValueError):
            Bridge(self.s).handle('save_settings',dict(weekdays='0,1,2,3,4,5',start='09:00',end='18:00',morning='09:15',evening='17:30',holidays='',catchup_enabled='1',catchup_hours='169'))
    def test_uncertain_send_is_not_retried(self):
        r=self.enroll();(self.s.home/'STOP').unlink();self.s.claim(r['id'],self.now+dt.timedelta(minutes=6))
        self.s.fail(r['id'],1,'Connection dropped after Send')
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(minutes=7)))
        self.s.resolve(r['id']);self.assertEqual(self.s.get(r['id'])['count'],1)
        self.assertEqual(self.s.get(r['id'])['status'],'Paused')
    def test_stale_editor_cannot_override_reply(self):
        r=self.enroll();self.s.action(r['id'],'Pause')
        with self.assertRaises(ValueError):self.s.edit(r['id'],self.fields,r['version'],self.now)
    def test_nested_transaction_rolls_back(self):
        with self.assertRaises(RuntimeError):
            with self.s.tx():
                self.enroll()
                raise RuntimeError('abort')
        self.assertEqual(self.s.all(),[])
    def test_workbook_import_atomic_and_no_resume(self):
        import xml.etree.ElementTree as ET
        r=self.enroll();self.s.action(r['id'],'Pause');r=self.s.get(r['id'])
        root=ET.Element('rows')
        for version in (r['version'],r['version']+10):
            node=ET.SubElement(root,'row')
            for k,v in r.items():
                if not isinstance(v,dict):ET.SubElement(node,k).text=str(v)
            node.find('version').text=str(version)
        with self.assertRaises(ValueError):Bridge(self.s).handle('import',{'batch':ET.tostring(root).decode()})
        self.assertEqual(self.s.get(r['id'])['version'],r['version'])
    def test_stop_date_prevents_slot_after_end(self):
        self.fields.update(next_at='2026-09-14T19:00:00+05:30',stop_date='2026-09-14')
        with self.assertRaises(ValueError):self.enroll()
    def test_generated_draft_permit_rejected_after_close(self):
        r=self.enroll();(self.s.home/'STOP').unlink();self.s.claim(r['id'],self.now+dt.timedelta(minutes=6))
        self.s.action(r['id'],'Close')
        response=Bridge(self.s).handle('permit',{'id':r['id'],'occ':'1'})
        self.assertIn(b'<allowed>0</allowed>',response)
    def test_review_hours_cross_weekend(self):
        prefs=dict(DEFAULTS,weekdays=[0,1,2,3,4])
        result=working_hours(dt.datetime(2026,9,11,17,tzinfo=IST),2,prefs)
        self.assertEqual(result,dt.datetime(2026,9,14,10,tzinfo=IST))
    def test_concurrent_claims_have_one_winner(self):
        from concurrent.futures import ThreadPoolExecutor
        r=self.enroll();(self.s.home/'STOP').unlink()
        with ThreadPoolExecutor(max_workers=8) as pool:
            results=list(pool.map(lambda _:self.s.claim(r['id'],self.now+dt.timedelta(minutes=6)),range(16)))
        self.assertEqual(sum(x is not None for x in results),1)
    def test_auto_mode_requires_global_enable(self):
        self.fields['mode']='Auto';r=self.enroll();(self.s.home/'STOP').unlink()
        self.assertIsNone(self.s.claim(r['id'],self.now+dt.timedelta(minutes=6)))
    def test_old_imported_reply_does_not_pause(self):
        r=self.enroll()
        self.s.ingest(dict(entry='old',store='store1',conv='thread1',sent='0',received=stamp(self.now-dt.timedelta(days=2))))
        self.assertEqual(self.s.get(r['id'])['status'],'Active')
    def test_burst_ceiling_latches_stop(self):
        r=self.enroll();(self.s.home/'STOP').unlink()
        with self.s.tx():
            for n in range(5):self.assertTrue(self.s.dispatch_allowed(r['id'],n+1))
            self.assertFalse(self.s.dispatch_allowed(r['id'],6))
        self.assertTrue(self.s.blocked())
    def test_bulk_action_is_atomic_on_invalid_id(self):
        r=self.enroll()
        with self.assertRaises(ValueError):Bridge(self.s).handle('bulk',dict(ids=r['id']+'\nmissing',method='Close'))
        self.assertEqual(self.s.get(r['id'])['status'],'Active')

    def test_remove_hides_followup_and_never_reopens_on_reply_or_failure(self):
        r=self.enroll();fid=r['id']
        self.s.action(fid,'Remove')
        self.assertEqual(self.s.get(fid)['status'],'Removed')
        self.assertNotIn(fid,Bridge(self.s).handle('list',{}).decode())
        reply=dict(entry='removed-reply',store='store1',conv='thread1',mid='<removed-reply>',sent='0',received=stamp(self.now+dt.timedelta(minutes=1)),subject='Re: Test')
        self.s.ingest(reply)
        self.assertEqual(self.s.get(fid)['status'],'Removed')
        with self.s.tx():
            self.s.db.execute('INSERT INTO occurrences VALUES(?,?,?,?,?)',(fid,1,'Claimed',stamp(self.now),'{}'))
        self.s.fail(fid,1,'late failure')
        self.assertEqual(self.s.get(fid)['status'],'Removed')

    def test_removed_restore_is_paused_and_purge_requires_removed(self):
        r=self.enroll();fid=r['id']
        with self.assertRaises(ValueError):self.s.restore_removed(fid)
        with self.assertRaises(ValueError):self.s.purge_removed(fid)
        self.s.action(fid,'Remove')
        restored=self.s.restore_removed(fid)
        self.assertEqual(restored['status'],'Paused')
        self.assertIsNone(self.s.claim(fid,self.now+dt.timedelta(days=2)))
        self.s.action(fid,'Remove');self.s.purge_removed(fid)
        with self.assertRaises(ValueError):self.s.get(fid)

    def test_dashboard_counts_and_removed_visibility_are_explicit(self):
        active=self.enroll()
        other=dict(self.anchor,entry='removed-original',conv='thread-removed',mid='<removed@example.com>')
        removed=self.s.enroll(self.fields,other,self.now)
        self.s.action(removed['id'],'Remove')
        with patch('service.utcnow',return_value=self.now):
            default=Bridge(self.s).handle('list',{}).decode()
            complete=Bridge(self.s).handle('list',{'include_removed':'1'}).decode()
        self.assertNotIn(removed['id'],default)
        self.assertIn(removed['id'],complete)
        self.assertIn('<removed>1</removed>',complete)
        self.assertIn('<due_today>1</due_today>',complete)
    def test_late_send_error_cannot_reopen_closed_followup(self):
        r=self.enroll();(self.s.home/'STOP').unlink();self.s.claim(r['id'],self.now+dt.timedelta(minutes=6))
        self.s.action(r['id'],'Close');self.s.fail(r['id'],1,'COM failure after Close')
        self.assertEqual(self.s.get(r['id'])['status'],'Closed')

if __name__=='__main__':unittest.main()
