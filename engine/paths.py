"""Shared runtime data location. Keep VBA FOHome and Stop-Automation.cmd aligned."""
import os
from pathlib import Path

def data_home():
    return Path(r'D:\FollowupOrganizer\data')

def legacy_home():
    previous=Path(r'D:\FollowupOrganizer')
    if (previous/'organizer.sqlite3').is_file():return previous
    return Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'FollowupOrganizer'

def require_migration():
    old=legacy_home();new=data_home()
    if old.resolve()!=new.resolve() and (old/'organizer.sqlite3').is_file() and not (new/'organizer.sqlite3').exists():
        raise RuntimeError('Existing reminder data found in the old location. Close organizer processes and run Migrate-Existing-Data.cmd before starting this version.')
