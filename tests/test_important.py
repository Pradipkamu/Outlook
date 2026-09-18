import datetime as dt
import tempfile
import unittest
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from core import Store,IST,DEFAULT_IMPORTANT_CATEGORIES
from service import Bridge

class ImportantMailTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.s=Store(Path(self.tmp.name))
        self.message=dict(entry='e1',store='s1',mid='<one@example>',conv='c1',subject='Quality alert',
                          received='2026-09-18T10:00:00+05:30',folder='\\Inbox')
        self.fields=dict(category='Quality',priority='High',tags='customer;urgent',note='Check CAPA',review_date='2026-09-18')
    def tearDown(self):self.s.db.close();self.tmp.cleanup()
    def test_defaults_custom_category_and_duplicate_protection(self):
        self.assertEqual(self.s.meta('important_categories'),DEFAULT_IMPORTANT_CATEGORIES)
        saved=self.s.important_add(self.fields,self.message)
        self.assertEqual(saved['category'],'Quality')
        with self.assertRaises(ValueError):self.s.important_add(self.fields,self.message)
        second=dict(self.message,entry='e2',mid='<two@example>',conv='c2')
        fields=dict(self.fields,category='Safety')
        self.s.important_add(fields,second)
        self.assertIn('Safety',self.s.meta('important_categories'))
    def test_edit_archive_relink_and_delete(self):
        saved=self.s.important_add(self.fields,self.message)
        changed=self.s.important_edit(saved['id'],dict(self.fields,priority='Critical'),1)
        self.assertEqual(changed['priority'],'Critical')
        self.s.important_action(saved['id'],'Archive');self.assertEqual(self.s.important_get(saved['id'])['status'],'Archived')
        self.s.important_action(saved['id'],'Archive');self.assertEqual(self.s.important_get(saved['id'])['status'],'Saved')
        relink=self.s.important_relink(saved['id'],dict(self.message,entry='new',mid='<new@example>'))
        self.assertEqual(relink['entry'],'new');self.assertEqual(relink['status'],'Saved')
        self.s.important_action(saved['id'],'Delete')
        with self.assertRaises(ValueError):self.s.important_get(saved['id'])
    def test_bridge_formats_ist_and_marks_review_due(self):
        self.s.important_add(self.fields,self.message,dt.datetime(2026,9,18,tzinfo=IST))
        xml=Bridge(self.s).handle('important_list',{}).decode()
        self.assertIn('<received_local>18 Sep 2026 10:00 IST</received_local>',xml)
        self.assertIn('<review_due>1</review_due>',xml)
    def test_validation_limits(self):
        with self.assertRaises(ValueError):self.s.important_add(dict(self.fields,category=''),self.message)
        with self.assertRaises(ValueError):self.s.important_add(dict(self.fields,priority='Emergency'),self.message)
        with self.assertRaises(ValueError):self.s.important_add(dict(self.fields,review_date='18/09/2026'),self.message)

    def test_approved_related_mail_is_grouped_without_followup(self):
        saved=self.s.important_add(self.fields,self.message)
        related=dict(self.message,entry='e2',mid='<related@example>',subject='RE: Quality alert')
        result=self.s.important_attach(saved['id'],related)
        self.assertTrue(result['linked']);self.assertEqual(result['count'],2)
        self.assertEqual(len(self.s.important_get(saved['id'])['links']),2)
        self.assertFalse(self.s.important_attach(saved['id'],related)['linked'])

if __name__=='__main__':unittest.main()
