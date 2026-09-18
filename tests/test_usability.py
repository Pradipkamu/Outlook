import datetime as dt
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from core import Store, IST, stamp
from service import Bridge

class UsabilityTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.home=Path(self.tmp.name);self.s=Store(self.home)
        self.now=dt.datetime(2026,9,14,9,tzinfo=IST)
        fields=dict(subject='Test',to='team@example.com',account='me@example.com',next_at=stamp(self.now+dt.timedelta(minutes=5)),every='1',unit='working days',maximum='5',mode='Draft',message='Update please')
        self.original=dict(entry='original',store='s',conv='thread',mid='<original>',sent='1')
        self.r=self.s.enroll(fields,self.original,self.now);self.fid=self.r['id']
        self.reply=dict(entry='reply',store='s',conv='thread',mid='<reply>',sent='0',fid=self.fid,received=stamp(self.now+dt.timedelta(minutes=1)))
        self.s.ingest(self.reply)
    def tearDown(self):self.s.db.close();self.tmp.cleanup()
    def unlink(self,m,**kw):return self.s.unlink(self.fid,m['entry'],m['store'],self.s.get(self.fid)['version'],**kw)
    def test_unlink_reply_preserves_review_schedule_and_stop(self):
        before=self.s.get(self.fid);blocked=self.s.blocked();r=self.unlink(self.reply)
        self.assertEqual(r['status'],'Review');self.assertEqual(r['next_at'],before['next_at']);self.assertEqual(blocked,self.s.blocked())
        self.assertEqual(len(self.s.messages(self.fid)),1)
        self.assertEqual(self.s.ingest(self.reply),'')
    def test_exclusion_survives_restart_and_message_move(self):
        self.unlink(self.reply);self.s.db.close();self.s=Store(self.home)
        moved=dict(self.reply,entry='moved',fid='')
        self.assertEqual(self.s.ingest(moved),'')
        no_mid=dict(self.reply,mid='');self.assertEqual(self.s.ingest(no_mid),'')
    def test_explicit_relink_clears_exclusion(self):
        self.unlink(self.reply);self.s.manual_link(self.fid,self.reply)
        self.assertFalse(self.s.excluded(self.reply));self.assertEqual(self.s.ingest(self.reply),self.fid)
    def test_anchor_requires_valid_replacement_and_rolls_back(self):
        for kwargs in ({},dict(replacement_entry='missing',replacement_store='s'),dict(replacement_entry='original',replacement_store='s')):
            with self.assertRaises(ValueError):self.unlink(self.original,**kwargs)
            self.assertEqual(len(self.s.messages(self.fid)),2);self.assertFalse(self.s.excluded(self.original))
        r=self.unlink(self.original,replacement_entry='reply',replacement_store='s')
        self.assertEqual(r['anchor']['entry'],'reply');self.assertEqual(r['status'],'Review')
        self.assertEqual(self.s.ingest(self.original),'')
    def test_last_email_cannot_be_removed(self):
        self.unlink(self.reply)
        with self.assertRaises(ValueError):self.unlink(self.original)
        self.assertEqual(len(self.s.messages(self.fid)),1)
    def test_stale_unlink_cannot_override_reply(self):
        with self.assertRaises(ValueError):self.s.unlink(self.fid,'reply','s',self.r['version'])
        self.assertEqual(len(self.s.messages(self.fid)),2)
    def test_generated_history_and_outstanding_send_protected(self):
        generated=dict(entry='generated',store='s',mid='<generated>',generated='1',occ='1')
        with self.s.tx():self.s.link(self.fid,generated)
        with self.assertRaises(ValueError):self.unlink(generated)
        with self.assertRaises(ValueError):self.unlink(self.original,replacement_entry='generated',replacement_store='s')
        for status in ('Draft','Claimed','Uncertain'):
            with self.s.tx():self.s.db.execute('INSERT OR REPLACE INTO occurrences VALUES(?,?,?,?,?)',(self.fid,1,status,stamp(self.now),'{}'))
            with self.assertRaises(ValueError):self.unlink(self.reply)
            with self.assertRaises(ValueError):self.s.manual_link(self.fid,self.reply)
        self.assertEqual(len(self.s.messages(self.fid)),3)
    def test_unlink_endpoint_and_explicit_link(self):
        bridge=Bridge(self.s)
        bridge.handle('unlink',dict(id=self.fid,entry='reply',store='s',version=str(self.s.get(self.fid)['version'])))
        self.assertEqual(self.s.ingest(self.reply),'')
        bridge.handle('link',dict(id=self.fid,**{'a_'+k:v for k,v in self.reply.items()}))
        self.assertEqual(self.s.ingest(self.reply),self.fid)

    def test_related_attachment_links_without_changing_reply_source(self):
        anchor=dict(self.s.get(self.fid)['anchor']);version=self.s.get(self.fid)['version']
        related=dict(entry='related',store='s',conv='other',mid='<related>',sent='0',subject='RE: Test',received=stamp(self.now))
        result=self.s.attach_related(self.fid,related)
        self.assertTrue(result['linked'])
        self.assertEqual(result['anchor'],anchor)
        self.assertEqual(result['status'],'Review')
        self.assertEqual(result['version'],version+1)
        self.assertEqual(self.s.attach_related(self.fid,related)['linked'],False)
    def test_history_endpoint_displays_ist_not_stored_utc(self):
        self.s.db.execute("UPDATE events SET at=? WHERE fid=? AND action='Reply received'",(stamp(self.now+dt.timedelta(minutes=1)),self.fid))
        self.s.db.commit()
        xml=Bridge(self.s).handle('history',{'id':self.fid}).decode()
        self.assertIn('<at>14 Sep 2026 09:01:00 IST</at>',xml)
        self.assertNotIn('+00:00',xml)

    def test_history_endpoint_is_newest_first_after_ist_formatting(self):
        self.s.log(self.fid,'Older event','')
        self.s.log(self.fid,'Newest event','')
        self.s.db.execute("UPDATE events SET at=? WHERE fid=?",(stamp(self.now),self.fid))
        self.s.db.execute("UPDATE events SET at=? WHERE fid=? AND action='Older event'",(stamp(self.now+dt.timedelta(hours=1)),self.fid))
        self.s.db.execute("UPDATE events SET at=? WHERE fid=? AND action='Newest event'",(stamp(self.now+dt.timedelta(hours=2)),self.fid))
        self.s.db.commit()
        rows=ET.fromstring(Bridge(self.s).handle('history',{'id':self.fid})).findall('row')
        timestamps=[dt.datetime.strptime(row.findtext('at'),'%d %b %Y %H:%M:%S IST') for row in rows]
        self.assertEqual(timestamps,sorted(timestamps,reverse=True))
        self.assertEqual(rows[0].findtext('action'),'Newest event')
    def test_conversation_endpoint_displays_aware_and_legacy_dates_in_ist(self):
        legacy=dict(entry='legacy',store='s',conv='thread',mid='<legacy>',received='2026-09-16T09:57:42')
        with self.s.tx():self.s.link(self.fid,legacy)
        xml=Bridge(self.s).handle('messages',{'id':self.fid}).decode()
        self.assertIn('<received>14 Sep 2026 09:01:00 IST</received>',xml)
        self.assertIn('<received>16 Sep 2026 09:57:42 IST</received>',xml)
        self.assertNotIn('+00:00',xml)

    def test_conversation_endpoint_is_newest_first_and_missing_dates_last(self):
        newer=dict(entry='newer',store='s',conv='thread',mid='<newer>',received=stamp(self.now+dt.timedelta(days=2)))
        undated=dict(entry='undated',store='s',conv='thread',mid='<undated>',received='')
        with self.s.tx():
            self.s.link(self.fid,newer)
            self.s.link(self.fid,undated)
        rows=ET.fromstring(Bridge(self.s).handle('messages',{'id':self.fid})).findall('row')
        entries=[row.findtext('entry') for row in rows]
        self.assertEqual(entries[:2],['newer','reply'])
        self.assertCountEqual(entries[2:],['original','undated'])
    def test_appearance_is_persistent_separate_and_validated_atomically(self):
        bridge=Bridge(self.s);settings=self.s.meta('settings');blocked=self.s.blocked()
        self.assertIn(b'<dash_font>Tahoma</dash_font>',bridge.handle('appearance',{}))
        bridge.handle('save_appearance',dict(dash_font='Arial',dash_size='10',dash_color='Navy',edit_color='Purple'))
        saved=self.s.meta('appearance')
        for values in (dict(dash_font='Unknown'),dict(edit_size='40'),dict(edit_color='White')):
            with self.assertRaises(ValueError):bridge.handle('save_appearance',values)
            self.assertEqual(self.s.meta('appearance'),saved)
        self.s.db.close();self.s=Store(self.home)
        self.assertEqual(self.s.meta('appearance'),saved);self.assertEqual(self.s.meta('settings'),settings);self.assertEqual(self.s.blocked(),blocked)

if __name__=='__main__':unittest.main()
