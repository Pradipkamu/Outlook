import sys
import subprocess
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace,ModuleType
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from startup import install,remove
from background_start import launch
from outlook_worker import OutlookWorker

class StartupTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.base=Path(self.temp.name)
        self.root=self.base/'Folder with spaces & marks';self.folder=self.base/'Startup';self.folder.mkdir()
        for name in ('.venv/Scripts/python.exe','.venv/Scripts/pythonw.exe','engine/background_start.py','engine/service.py','engine/watchdog.py'):
            p=self.root/name;p.parent.mkdir(parents=True,exist_ok=True);p.touch()
    def tearDown(self):self.temp.cleanup()
    def test_install_is_repeatable_and_quotes_script(self):
        links=[]
        def create(path):
            link=SimpleNamespace(Save=lambda:Path(path).touch());links.append(link);return link
        shell=SimpleNamespace(CreateShortcut=create)
        install(self.root,self.folder,shell);install(self.root,self.folder,shell)
        self.assertEqual(len(list(self.folder.iterdir())),1)
        # Windows TEMP may use an 8.3 alias which install().resolve() expands.
        # Compare file identity, not the spelling/case of equivalent paths.
        for link in links:
            self.assertTrue(Path(link.TargetPath).samefile(self.root/'.venv/Scripts/pythonw.exe'))
            self.assertTrue(Path(link.WorkingDirectory).samefile(self.root))
            expected_script=self.root.resolve()/'engine'/'background_start.py'
            self.assertEqual(link.Arguments,subprocess.list2cmdline([str(expected_script)]))
            self.assertTrue(link.Arguments.startswith('"'))
    def test_missing_environment_does_not_create_shortcut(self):
        (self.root/'.venv/Scripts/pythonw.exe').unlink()
        with self.assertRaises(FileNotFoundError):install(self.root,self.folder,None)
        self.assertEqual(list(self.folder.iterdir()),[])
    def test_remove_leaves_other_startup_entries(self):
        (self.folder/'Follow-up Organizer.lnk').touch();(self.folder/'Other.lnk').touch()
        remove(self.folder);remove(self.folder)
        self.assertTrue((self.folder/'Other.lnk').exists())
    def test_background_launch_preserves_safety_stop_and_uses_argument_arrays(self):
        home=self.base/'data';home.mkdir();(home/'STOP').write_text('Keep stopped')
        calls=[]
        launch(self.root,home,lambda args,**kwargs:calls.append((args,kwargs)))
        self.assertEqual(len(calls),2)
        self.assertEqual(Path(calls[0][0][1]).name,'service.py')
        self.assertEqual(Path(calls[1][0][1]).name,'watchdog.py')
        self.assertFalse(calls[0][1].get('shell',False))
        self.assertEqual((home/'STOP').read_text(),'Keep stopped')
    def test_absent_outlook_is_waiting_but_other_com_errors_propagate(self):
        class COMError(Exception):
            def __init__(self,code):self.hresult=code
        win=ModuleType('win32com');client=ModuleType('win32com.client');win.client=client
        pythoncom=ModuleType('pythoncom');pythoncom.com_error=COMError
        def unavailable(*args):raise COMError(-2147221021)
        client.GetActiveObject=unavailable
        with patch.dict(sys.modules,{'win32com':win,'win32com.client':client,'pythoncom':pythoncom}):
            self.assertFalse(OutlookWorker(None).connect())
            def denied(*args):raise COMError(-2147024891)
            client.GetActiveObject=denied
            with self.assertRaises(COMError):OutlookWorker(None).connect()

if __name__=='__main__':unittest.main()
