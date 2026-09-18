"""Install/remove this user's Startup shortcut. No elevation or policy changes."""
import argparse
import os
import subprocess
from pathlib import Path
from paths import data_home,require_migration

def install(root,folder,shell):
    root=Path(root).resolve();folder=Path(folder)
    target=root/'.venv'/'Scripts'/'pythonw.exe';script=root/'engine'/'background_start.py'
    if not target.is_file() or not script.is_file():raise FileNotFoundError('Run Setup.cmd in the complete extracted package first.')
    shortcut_path=folder/'Follow-up Organizer.lnk'
    shortcut=shell.CreateShortcut(str(shortcut_path))
    shortcut.TargetPath=str(target)
    shortcut.Arguments=subprocess.list2cmdline([str(script)])
    shortcut.WorkingDirectory=str(root)
    shortcut.Description='Follow-up Organizer and independent watchdog; current user sign-in'
    shortcut.WindowStyle=7
    shortcut.Save()
    if not shortcut_path.is_file():raise OSError('Windows did not save the startup shortcut.')
    return shortcut_path

def remove(folder):
    (Path(folder)/'Follow-up Organizer.lnk').unlink(missing_ok=True)

def main():
    parser=argparse.ArgumentParser();parser.add_argument('action',choices=['install','remove']);args=parser.parse_args()
    import win32com.client
    from win32com.shell import shell,shellcon
    folder=Path(shell.SHGetFolderPath(0,shellcon.CSIDL_STARTUP,None,0))
    root=Path(__file__).resolve().parents[1]
    if args.action=='remove':
        remove(folder)
        print('Automatic startup removed for this Windows user. Running processes and reminder data are unchanged.')
        return
    require_migration()
    install(root,folder,win32com.client.Dispatch('WScript.Shell'))
    print('Automatic startup installed for this Windows user.')
    from background_start import launch
    launch(root,data_home())
    print('Background launch requested now. Open Classic Outlook normally. Existing safety stops remain in force.')
    print('After your next sign-in, you do not need to run either organizer CMD manually.')

if __name__=='__main__':main()
