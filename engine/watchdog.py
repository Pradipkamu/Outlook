"""Independent fail-closed watchdog. Never kills Outlook or clears a stop."""
import json
import os
import time
from pathlib import Path
from paths import data_home
home=data_home()
home.mkdir(parents=True,exist_ok=True)
import msvcrt
lockfile=open(home/'watchdog.lock','a+b');lockfile.seek(0);lockfile.write(b'1');lockfile.flush();lockfile.seek(0)
try:msvcrt.locking(lockfile.fileno(),msvcrt.LK_NBLCK,1)
except OSError:raise SystemExit('Watchdog already running.')
started=time.monotonic()
while True:
    try:
        hb=json.loads((home/'heartbeat.json').read_text())
        stale=time.time()-hb['time']>45
    except (OSError,ValueError,KeyError):stale=True
    if stale and time.monotonic()-started>=45 and not (home/'STOP').exists():
        (home/'STOP').write_text('Watchdog: processing heartbeat missing for 45 seconds. Check engine and review before resuming.',encoding='utf-8')
    time.sleep(5)
