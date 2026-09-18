import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
import paths
from migrate_data import copy_data

class DataLocationTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.old=Path(self.temp.name)/'old';self.new=Path(self.temp.name)/'new'
        self.old.mkdir();self.new.mkdir()
        db=sqlite3.connect(self.old/'organizer.sqlite3');db.execute('CREATE TABLE test(value TEXT)');db.execute("INSERT INTO test VALUES ('keep reminder')");db.commit();db.close()
        (self.old/'token.txt').write_text('existing-local-token')
        (self.old/'STOP').write_text('User stopped sending')
    def tearDown(self):self.temp.cleanup()
    def test_exact_requested_location(self):
        self.assertEqual(str(paths.data_home()),r'D:\FollowupOrganizer\data')
    def test_migration_preserves_records_token_and_stop(self):
        (self.old/'worker.log').write_text('old log');(self.new/'worker.log').write_text('new log')
        copy_data(self.old,self.new)
        db=sqlite3.connect(self.new/'organizer.sqlite3')
        try:self.assertEqual(db.execute('SELECT value FROM test').fetchone()[0],'keep reminder')
        finally:db.close()
        self.assertEqual((self.new/'token.txt').read_text(),'existing-local-token')
        self.assertIn('User stopped sending',(self.new/'STOP').read_text())
        self.assertEqual((self.new/'worker.log').read_text(),'new log')
        self.assertEqual((self.new/'MigratedLogs/worker.log').read_text(),'old log')
        self.assertTrue((self.old/'organizer.sqlite3').exists())
    def test_existing_destination_database_not_overwritten(self):
        (self.new/'organizer.sqlite3').write_bytes(b'keep this')
        with self.assertRaises(FileExistsError):copy_data(self.old,self.new)
        self.assertEqual((self.new/'organizer.sqlite3').read_bytes(),b'keep this')
    def test_start_blocked_until_legacy_data_migrated(self):
        with patch.object(paths,'legacy_home',return_value=self.old),patch.object(paths,'data_home',return_value=self.new):
            with self.assertRaises(RuntimeError):paths.require_migration()
            copy_data(self.old,self.new);paths.require_migration()

if __name__=='__main__':unittest.main()
