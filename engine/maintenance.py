"""Daily snapshots and metadata-only diagnostics; no raw mail or log export."""
import datetime as dt
import json
import os
import platform
import re
import sqlite3
import tempfile
import time
import zipfile
from collections import Counter
from pathlib import Path
from core import IST,utcnow,stamp
from version import VERSION,BUILD,RELEASE

def backup_snapshot(home,daily=True,now=None):
    home=Path(home);now=now or utcnow();dest=home/'Backups';dest.mkdir(exist_ok=True)
    name=('Auto-'+now.astimezone(IST).strftime('%Y%m%d') if daily else 'Manual-'+now.astimezone(IST).strftime('%Y%m%d-%H%M%S-%f'))+'.zip'
    out=dest/name
    if daily and out.exists():
        try:
            with zipfile.ZipFile(out) as z:
                if z.testzip() is None and 'organizer.sqlite3' in z.namelist():return out
        except (OSError,zipfile.BadZipFile):pass
    with tempfile.TemporaryDirectory(prefix='.backup-',dir=dest) as tmp:
        database=Path(tmp)/'organizer.sqlite3'
        source=sqlite3.connect((home/'organizer.sqlite3').resolve().as_uri()+'?mode=ro',uri=True)
        target=sqlite3.connect(database)
        try:
            source.backup(target,pages=256,sleep=0.01)
            if target.execute('PRAGMA integrity_check').fetchone()[0]!='ok':raise RuntimeError('Backup integrity check failed')
        finally:target.close();source.close()
        staged=Path(tmp)/'snapshot.zip'
        with zipfile.ZipFile(staged,'w',zipfile.ZIP_DEFLATED) as z:
            z.write(database,'organizer.sqlite3')
            z.writestr('STOP','Restored backup: keep stopped until pending sends and replies are reconciled.')
            z.writestr('README.txt','Close organizer processes before restoring. Preserve current data first. Snapshot includes configuration/history and message metadata; keep private. It excludes token.txt. Resume only after reviewing Outlook and uncertain sends.')
        with zipfile.ZipFile(staged) as z:
            if z.testzip() is not None:raise RuntimeError('Backup archive verification failed')
        os.replace(staged,out)
    # Rotate ONLY our daily snapshots. Manual/legacy backups are never removed.
    daily_files=sorted(p for p in dest.iterdir() if re.fullmatch(r'Auto-\d{8}\.zip',p.name))
    for old in daily_files[:-30]:old.unlink()
    return out

def verify_backup(path):
    """Verify archive structure and the SQLite snapshot without restoring it."""
    path=Path(path)
    if not path.is_file():raise ValueError('Backup file does not exist.')
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None:raise RuntimeError('Backup archive contains damaged data.')
        names=set(z.namelist())
        if 'organizer.sqlite3' not in names or not {'STOP','README.txt'}.issubset(names):
            raise RuntimeError('Backup archive is incomplete.')
        if any(name.startswith(('/', '\\')) or '..' in Path(name).parts for name in names):
            raise RuntimeError('Backup archive contains an unsafe path.')
        with tempfile.TemporaryDirectory(prefix='.verify-',dir=path.parent) as tmp:
            database=Path(tmp)/'organizer.sqlite3'
            with z.open('organizer.sqlite3') as source, database.open('wb') as target:
                while True:
                    block=source.read(1024*1024)
                    if not block:break
                    target.write(block)
            with sqlite3.connect(database) as db:
                if db.execute('PRAGMA integrity_check').fetchone()[0]!='ok':
                    raise RuntimeError('Backup database integrity check failed.')
                reminders=db.execute('SELECT COUNT(*) FROM reminders').fetchone()[0]
                important=db.execute('SELECT COUNT(*) FROM important_mails').fetchone()[0]
    return dict(file=path.name,size=path.stat().st_size,reminders=reminders,important=important,verified_at=stamp(utcnow()))

def latest_backup(home):
    files=sorted((Path(home)/'Backups').glob('*.zip'),key=lambda p:p.stat().st_mtime,reverse=True)
    if not files:raise ValueError('No Organizer backup is available.')
    return files[0]

def maintenance_loop(store):
    retry_at=0
    while True:
        if time.monotonic()>=retry_at:
            try:
                path=backup_snapshot(store.home)
                state=store.meta('backup') or {}
                if state.get('file')!=path.name or state.get('error'):
                    store.setmeta('backup',dict(file=path.name,at=stamp(utcnow()),error=''))
                retry_at=time.monotonic()+60
            except Exception:
                # No paths or mail-derived exception strings in diagnostic metadata.
                store.setmeta('backup',dict(error='Backup failed: check disk space and data-folder permissions.',at=stamp(utcnow())))
                retry_at=time.monotonic()+3600
        time.sleep(60)

def health_summary(store):
    good,hb=store.health();stop=store.blocked();vba=store.home/'vba-heartbeat'
    vba_ok=vba.exists() and 0<=time.time()-vba.stat().st_mtime<180
    detail=hb.get('detail','')
    if stop:
        state='Paused by you' if stop.startswith(('User pressed','User requested')) else 'Safety stop - review required'
    elif not good:state='Error - engine heartbeat missing'
    elif 'Waiting for Classic Outlook' in detail:state='Waiting for Outlook'
    elif 'Waiting for before-send' in detail:state='Waiting for your settings'
    elif not vba_ok:state='Error - VBA monitor stale'
    elif 'failed' in detail.lower() or 'unavailable' in detail.lower():state='Error - Outlook connection'
    elif hb.get('ready'):state='Running'
    else:state='Checking Outlook'
    return dict(state=state,vba_ok=vba_ok,heartbeat_ok=good,last_scan=store.meta('last_successful_scan') or '',backup=store.meta('backup') or {})

def export_diagnostics(store):
    """Allowlist only. Raw exception messages may contain recipients/subjects/paths."""
    health=health_summary(store)
    with store.lock:
        states=Counter(r['status'] for r in store.all())
        rows=store.db.execute("SELECT at,action FROM events WHERE action IN ('Send outcome needs review','Dispatch','Submitted','Pause','Close','Resume') ORDER BY seq DESC LIMIT 50").fetchall()
    report=dict(version=VERSION,build=BUILD,release=RELEASE,python=platform.python_version(),os=platform.system(),
                generated_at=stamp(utcnow()),health={k:v for k,v in health.items() if k!='backup'},
                backup_ok=bool(health['backup'].get('file')) and not bool(health['backup'].get('error')),
                counts=dict(states),recent_events=[dict(r) for r in rows])
    dest=store.home/'Diagnostics';dest.mkdir(exist_ok=True)
    path=dest/('Organizer-diagnostics-'+utcnow().strftime('%Y%m%d-%H%M%S-%f')+'.json')
    path.write_text(json.dumps(report,indent=2),encoding='utf-8')
    return path
