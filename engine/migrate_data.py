"""Copy legacy runtime data, with both process locations locked and no overwrites."""
import contextlib
import shutil
import sqlite3
import tempfile
from pathlib import Path
from paths import data_home,legacy_home

@contextlib.contextmanager
def process_locks(folders):
    import msvcrt
    handles=[]
    try:
        for folder in folders:
            folder.mkdir(parents=True,exist_ok=True)
            for name in ('service.lock','watchdog.lock'):
                handle=open(folder/name,'a+b');handles.append(handle)
                if (folder/name).stat().st_size==0:
                    handle.write(b'1');handle.flush()
                handle.seek(0)
                try:msvcrt.locking(handle.fileno(),msvcrt.LK_NBLCK,1)
                except OSError:raise RuntimeError('An organizer/watchdog is still running. Remove automatic startup, close Outlook, restart Windows, and run migration before starting the organizer again.')
        yield
    finally:
        for handle in reversed(handles):handle.close()

def copy_data(old,new):
    """Caller must hold process_locks; leaves all original files intact."""
    old=Path(old);new=Path(new)
    if old.resolve()==new.resolve():return 'Data is already in the requested location.'
    db=old/'organizer.sqlite3'
    if not db.is_file():return 'No existing reminder database found; a new one will be created on first start.'
    new.mkdir(parents=True,exist_ok=True)
    # Refuse even an empty new DB: do not merge two independent histories.
    if (new/'organizer.sqlite3').exists():raise FileExistsError('Destination already contains a database. No data overwritten; review both locations before merging.')
    files=[p for p in old.iterdir() if p.is_file() and (p.name=='token.txt' or p.suffix=='.log' or (p.name.startswith('Followups-') and p.suffix=='.xlsx'))]
    for name in ('Backups','MigratedLogs','Diagnostics'):
        folder=old/name
        if folder.is_dir():files.extend(p for p in folder.rglob('*') if p.is_file())
    # Root logs may share names with logs retained by an earlier migration.
    # Reserve nested paths first, then choose distinct names for current root logs.
    destinations={p:new/p.relative_to(old) for p in files if not (p.suffix=='.log' and p.parent==old)}
    reserved=set(destinations.values())
    for p in files:
        if p in destinations:continue
        dest=new/'MigratedLogs'/p.name
        for n in range(len(files)+1):
            if dest not in reserved:break
            dest=new/'MigratedLogs'/('previous-root-'+str(n)+'-'+p.name)
        destinations[p]=dest;reserved.add(dest)
    def destination(p):return destinations[p]
    for p in files:
        if destination(p).exists():raise FileExistsError('Destination file already exists: '+str(destination(p))+'. No data overwritten.')
    reason='Data location changed. Check migrated follow-ups and sent outcomes before Review and resume.'
    if (old/'STOP').exists():reason+=' Previous stop: '+(old/'STOP').read_text(encoding='utf-8')
    # Stage snapshot and check integrity before publishing any migrated file.
    with tempfile.TemporaryDirectory(prefix='fo-migration-',dir=new) as tmp:
        staged=Path(tmp)/'organizer.sqlite3'
        source=sqlite3.connect(db.resolve().as_uri()+'?mode=ro',uri=True)
        target=sqlite3.connect(staged)
        try:
            source.backup(target)
            if target.execute('PRAGMA integrity_check').fetchone()[0]!='ok':raise RuntimeError('Database integrity check failed.')
        finally:target.close();source.close()
        for p in files:
            dest=destination(p);dest.parent.mkdir(parents=True,exist_ok=True)
            with dest.open('xb') as output,p.open('rb') as source_file:shutil.copyfileobj(source_file,output)
        # New heartbeats and process locks are deliberately not copied.
        (new/'STOP').write_text(reason,encoding='utf-8')
        with (new/'organizer.sqlite3').open('xb') as output,staged.open('rb') as source_file:shutil.copyfileobj(source_file,output)
    return 'Reminder database, credential, logs, Excel exports and backups copied. Old data retained; new location is stopped for review.'

def main():
    old=legacy_home();new=data_home()
    if not (old/'organizer.sqlite3').exists():print(copy_data(old,new));return
    with process_locks([old,new]):print(copy_data(old,new))

if __name__=='__main__':main()
