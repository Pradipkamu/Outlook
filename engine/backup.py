"""Manual consistent backup; automatic daily backups run in the service."""
from paths import data_home,require_migration
from maintenance import backup_snapshot
if __name__ == '__main__':
    require_migration()
    print('Backup saved:',backup_snapshot(data_home(),daily=False))
