"""Start both existing single-instance processes without console windows."""
import os
import subprocess
import traceback
from pathlib import Path
from paths import data_home,require_migration

def launch(root,home,popen=subprocess.Popen):
    root=Path(root).resolve();home=Path(home)
    python=root/'.venv'/'Scripts'/'python.exe'
    scripts=[root/'engine'/'service.py',root/'engine'/'watchdog.py']
    for path in [python,*scripts]:
        if not path.is_file():raise FileNotFoundError(f'Missing {path}. Extract the complete package and run Setup.cmd.')
    home.mkdir(parents=True,exist_ok=True)
    for script in scripts:
        with (home/(script.stem+'-startup.log')).open('a',encoding='utf-8') as log:
            popen([str(python),str(script)],cwd=str(root),stdin=subprocess.DEVNULL,
                  stdout=log,stderr=subprocess.STDOUT,
                  creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    # Existing process locks reject duplicates. Never clear STOP or change Auto settings.

if __name__=='__main__':
    home=data_home()
    try:
        require_migration()
        launch(Path(__file__).resolve().parents[1],home)
    except Exception:
        home.mkdir(parents=True,exist_ok=True)
        with (home/'startup-error.log').open('a',encoding='utf-8') as log:traceback.print_exc(file=log)
        raise
