import datetime as dt
import json
import sqlite3
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from core import Store,IST,stamp,parse
from maintenance import backup_snapshot,export_diagnostics,health_summary,verify_backup,latest_backup
from service import Bridge
from migrate_data import copy_data

class ReleaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.home=Path(self.tmp.name);self.s=Store(self.home)
        self.now=dt.datetime(2026,9,14,9,tzinfo=IST)
        self.fields=dict(subject='PRIVATE SUBJECT',to='private@example.com',account='me@example.com',cc='',next_at=stamp(self.now+dt.timedelta(minutes=5)),every='1',unit='working days',maximum='5',mode='Draft',message='PRIVATE BODY',later_message='',priority='High',deadline='2026-09-16',stop_date='',note='PRIVATE NOTE')
        self.r=self.s.enroll(self.fields,dict(entry='original',store='store1',conv='thread1',mid='<private@example.com>',sent='1'),self.now)
    def tearDown(self):self.s.db.close();self.tmp.cleanup()
    def test_snooze_preserves_deadline_pause_and_stop(self):
        self.s.action(self.r['id'],'Pause');blocked=self.s.blocked()
        r=self.s.snooze(self.r['id'],'hour',self.now)
        self.assertEqual(parse(r['next_at']),self.now+dt.timedelta(hours=1))
        self.assertEqual(r['deadline'],'2026-09-16');self.assertEqual(r['status'],'Paused');self.assertEqual(self.s.blocked(),blocked)
    def test_snooze_cannot_skip_reply_review_or_outstanding_send(self):
        for status in ('Review','Closed','Pending send','Needs action','Limit reached','Draft ready'):
            self.r['status']=status;self.s.save(self.r)
            with self.assertRaises(ValueError):self.s.snooze(self.r['id'],'hour',self.now)
        self.r['status']='Active';self.s.save(self.r)
        for status in ('Draft','Claimed','Uncertain'):
            self.s.db.execute('INSERT OR REPLACE INTO occurrences VALUES(?,?,?,?,?)',(self.r['id'],1,status,stamp(self.now),'{}'));self.s.db.commit()
            with self.assertRaises(ValueError):self.s.snooze(self.r['id'],'hour',self.now)
    def test_snooze_workday_skips_weekend_and_holiday(self):
        prefs=self.s.meta('settings');prefs['weekdays']=[0,1,2,3,4];prefs['holidays']=['2026-09-21'];self.s.setmeta('settings',prefs)
        r=self.s.snooze(self.r['id'],'workday',dt.datetime(2026,9,18,17,tzinfo=IST))
        self.assertEqual(parse(r['next_at']),dt.datetime(2026,9,22,9,tzinfo=IST))
    def test_snooze_cannot_advance_schedule_or_cross_stop_date(self):
        self.r['stop_date']='2026-09-14';self.s.save(self.r)
        with self.assertRaises(ValueError):self.s.snooze(self.r['id'],'tomorrow',self.now)
        self.r['stop_date']='';self.r['next_at']=stamp(self.now+dt.timedelta(days=5));self.s.save(self.r)
        with self.assertRaises(ValueError):self.s.snooze(self.r['id'],'hour',self.now)
    def test_backup_snapshot_consistent_daily_and_private_token_excluded(self):
        (self.home/'token.txt').write_text('SECRET_TOKEN')
        path=backup_snapshot(self.home,now=self.now);before=path.stat().st_mtime_ns
        self.assertEqual(backup_snapshot(self.home,now=self.now),path);self.assertEqual(path.stat().st_mtime_ns,before)
        with zipfile.ZipFile(path) as z:
            self.assertNotIn('token.txt',z.namelist());self.assertIn('STOP',z.namelist());self.assertIsNone(z.testzip())
            z.extract('organizer.sqlite3',self.home/'check')
        with sqlite3.connect(self.home/'check/organizer.sqlite3') as db:self.assertEqual(db.execute('SELECT COUNT(*) FROM reminders').fetchone()[0],1)

    def test_manual_backup_can_be_located_and_fully_verified(self):
        path=backup_snapshot(self.home,daily=False,now=self.now)
        result=verify_backup(path)
        self.assertEqual(latest_backup(self.home),path)
        self.assertEqual(result['file'],path.name)
        self.assertEqual(result['reminders'],1)
        self.assertGreater(result['size'],0)
    def test_backup_retains_30_daily_and_manual(self):
        manual=backup_snapshot(self.home,daily=False,now=self.now)
        for n in range(32):backup_snapshot(self.home,now=self.now+dt.timedelta(days=n))
        self.assertEqual(len(list((self.home/'Backups').glob('Auto-*.zip'))),30);self.assertTrue(manual.exists())
    def test_backup_publish_failure_leaves_prior_snapshot(self):
        path=backup_snapshot(self.home,now=self.now)
        with patch('maintenance.os.replace',side_effect=OSError('disk full')):
            with self.assertRaises(OSError):backup_snapshot(self.home,now=self.now+dt.timedelta(days=1))
        self.assertTrue(path.exists());self.assertEqual(len(list((self.home/'Backups').glob('Auto-*.zip'))),1)
    def test_diagnostics_excludes_sensitive_values_and_raw_logs(self):
        self.s.log(self.r['id'],'Send outcome needs review','PRIVATE ERROR private@example.com')
        (self.home/'worker.log').write_text('PRIVATE LOG');(self.home/'token.txt').write_text('SECRET_TOKEN')
        result=export_diagnostics(self.s).read_text()
        for secret in ('PRIVATE','private@example.com','SECRET_TOKEN',self.r['id']):self.assertNotIn(secret,result)
        self.assertIn('Send outcome needs review',result)
    def test_health_stop_precedes_ready_and_waiting_is_clear(self):
        self.s.stop('User pressed Stop Automation.');(self.home/'vba-heartbeat').touch()
        with patch.object(self.s,'health',return_value=(True,dict(ready=True,detail='Scan complete'))):self.assertEqual(health_summary(self.s)['state'],'Paused by you')
        with patch.object(self.s,'blocked',return_value=''),patch.object(self.s,'health',return_value=(True,dict(ready=False,detail='Waiting for Classic Outlook'))):self.assertEqual(health_summary(self.s)['state'],'Waiting for Outlook')
    def test_nested_data_migration_preserves_program_and_history(self):
        self.s.db.close();(self.home/'program.py').write_text('keep program')
        (self.home/'MigratedLogs').mkdir();(self.home/'MigratedLogs/old.log').write_text('old')
        (self.home/'old.log').write_text('newer root log')
        copy_data(self.home,self.home/'data')
        with sqlite3.connect(self.home/'data/organizer.sqlite3') as db:self.assertEqual(db.execute('SELECT COUNT(*) FROM reminders').fetchone()[0],1)
        self.assertFalse((self.home/'data/program.py').exists());self.assertEqual((self.home/'program.py').read_text(),'keep program')
        self.assertEqual((self.home/'data/MigratedLogs/old.log').read_text(),'old')
        self.assertEqual((self.home/'data/MigratedLogs/previous-root-0-old.log').read_text(),'newer root log')
    def test_dashboard_distinguishes_submission_from_consumed_count(self):
        self.r['count']=2;self.s.save(self.r)
        self.s.db.execute('INSERT INTO occurrences VALUES(?,?,?,?,?)',(self.r['id'],1,'Submitted',stamp(self.now),'{}'))
        self.s.db.execute('INSERT INTO occurrences VALUES(?,?,?,?,?)',(self.r['id'],2,'Reviewed',stamp(self.now),'{}'));self.s.db.commit()
        self.s.log(self.r['id'],'Submitted','delivery not confirmed')
        xml=Bridge(self.s).handle('list',{}).decode()
        self.assertIn('<submitted_count>1</submitted_count>',xml);self.assertIn('<send_state>Reviewed</send_state>',xml)
    def test_contacts_dialog_uses_mutable_collection_without_readonly_assignment(self):
        code=(Path(__file__).resolve().parents[1]/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        self.assertIn('Set recipient = dialog.Recipients.Add(',code)
        self.assertNotIn('Set dialog.Recipients =',code)
        self.assertNotIn('seed.Recipients',code)

    def test_related_search_is_bounded_and_requires_per_message_approval(self):
        code=(Path(__file__).resolve().parents[1]/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        module=(Path(__file__).resolve().parents[1]/'vba'/'FollowupOrganizer.bas').read_text(encoding='utf-8')
        self.assertIn('RELATED_MAX_FOLDERS As Long = 200',code)
        self.assertIn('RELATED_MAX_MESSAGES As Long = 20000',code)
        self.assertIn('RELATED_MAX_CANDIDATES As Long = 100',code)
        self.assertIn('vbYesNoCancel + vbQuestion',code)
        self.assertIn('FOCall("attach_related"',code)
        self.assertIn('FOProp(mail, "FOGenerated") = ""',code)
        self.assertIn('Public Function FOInboxSnapshot',module)
        self.assertNotIn('FOIsOwn(mail)',module.split('Public Function FOInboxSnapshot',1)[1].split('End Function',1)[0])

    def test_remove_control_and_html_body_fix_are_present(self):
        form=(Path(__file__).resolve().parents[1]/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        worker=(Path(__file__).resolve().parents[1]/'engine'/'outlook_worker.py').read_text(encoding='utf-8')
        self.assertIn('Button "Remove", "Remove follow-up"',form)
        self.assertIn('Case "Pause", "Close", "Later", "Remove"',form)
        self.assertIn('mail.HTMLBody=reminder_html+original_html',worker)
        self.assertNotIn('mail.Body=message',worker)

    def test_summary_panel_is_transparent_borderless_and_selection_driven(self):
        root=Path(__file__).resolve().parents[1]
        form=(root/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        hook=(root/'vba'/'cFOListEvents.cls').read_text(encoding='utf-8')
        self.assertIn('c.BackStyle = fmBackStyleTransparent',form)
        self.assertIn('c.BorderStyle = fmBorderStyleNone',form)
        self.assertIn('Private Sub ShowSelectionSummary()',form)
        self.assertIn('Latest history:',form)
        self.assertIn('Private Sub List_Change()',hook)
        self.assertIn('Owner.SelectionChanged',hook)

    def test_v24_dashboard_recovery_and_backup_controls_are_present(self):
        root=Path(__file__).resolve().parents[1]
        form=(root/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        for action in ('CardOverdue','CardDue','CardAwaiting','CardReview','CardExceptions','CardRemoved','Restore','Purge','BackupNow','BackupCheck'):
            self.assertIn('"'+action+'"',form)
        self.assertIn('Type PURGE to continue:',form)
        self.assertIn('FOXML("include_removed", "1")',form)

    def test_v25_important_mail_register_is_present_and_separate(self):
        root=Path(__file__).resolve().parents[1]
        form=(root/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        module=(root/'vba'/'FollowupOrganizer.bas').read_text(encoding='utf-8')
        core=(root/'engine'/'core.py').read_text(encoding='utf-8')
        for action in ('ImportantSave','ImportantEdit','ImportantArchive','ImportantDelete','ImportantRelink','ImportantConvert'):
            self.assertIn('"'+action+'"',form)
        self.assertIn('Public Function FOImportantSnapshot',module)
        self.assertIn('CREATE TABLE IF NOT EXISTS important_mails',core)
        self.assertIn('DEFAULT_IMPORTANT_CATEGORIES',core)

    def test_v251_auto_link_searches_inbox_sent_and_asks_each_candidate(self):
        form=(Path(__file__).resolve().parents[1]/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        self.assertIn('GetDefaultFolder(olFolderInbox)',form)
        self.assertIn('GetDefaultFolder(olFolderSentMail)',form)
        self.assertIn('mail.ConversationID = targetConversation',form)
        self.assertIn('NormalizeRelatedSubject(mail.Subject) = targetSubject',form)
        self.assertIn('vbYesNoCancel + vbQuestion',form)
        self.assertIn('FOSnapshot(mail, "a_")',form)
        self.assertIn('SearchAndApproveImportantMail FOValue(result, "id"), mail',form)
        self.assertIn('FOCall("important_attach"',form)

    def test_v252_marks_mail_only_after_confirmed_link(self):
        form=(Path(__file__).resolve().parents[1]/'vba'/'Organizer-code.txt').read_text(encoding='utf-8')
        confirmation='linkedConfirmed = (LCase$(FOValue(result, "linked")) = "true")'
        self.assertIn(confirmation,form)
        confirmed=form.split(confirmation,1)[1].split('Else',1)[0]
        self.assertIn('If linkedConfirmed Then',confirmed)
        self.assertIn('MarkImportantMail mail, importantCategory',confirmed)

if __name__=='__main__':unittest.main()
