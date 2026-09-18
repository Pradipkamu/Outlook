# Data location — release 2.1

Follow **UPDATE-2.1.md** for the complete upgrade sequence.

Program files stay in `D:\FollowupOrganizer`. All application runtime files now use `D:\FollowupOrganizer\data`: organizer.sqlite3 and its journals, settings/history in that database, token.txt, STOP, heartbeats, process locks, logs, Excel exports, review markers, Backups, MigratedLogs and Diagnostics.

Migrate-Existing-Data.cmd first looks for organizer.sqlite3 directly in `D:\FollowupOrganizer`; if absent, it checks `%LOCALAPPDATA%\FollowupOrganizer`. It copies a consistent database snapshot, token, Excel exports, logs and backup/diagnostic folders. It leaves originals intact and sets a review stop. It does not copy program files, stale heartbeats or old lock contents. Original email messages and the non-reminding monitor task remain in Outlook.

No destination database or copied file is overwritten. If both locations contain databases, the tool refuses to guess or merge. An interrupted copy may leave partial files; retain both locations and share the error before retrying. Migration requires all organizer/watchdog processes to be stopped and takes Windows process locks to check this.

D: must be available and writable by your normal Windows user; there is no silent fallback. Only one Windows user/profile should operate this fixed data folder. Outlook's VBA project and Windows Startup shortcut remain in their normal system locations. Old runtime files remain outside data solely as a recovery copy until you verify migration; do not manually move a live SQLite database.
