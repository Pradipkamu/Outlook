# Update to Follow-up Organizer 2.4.0

## Daily management dashboard

Six live cards filter the working list:

- Overdue
- Due today
- Awaiting reply
- Review replies
- Exceptions
- Removed

Counts are calculated by the engine using IST-aware schedule dates.

## Exception Centre

The Exceptions card shows setup-required items, Needs action records, stored errors, uncertain send outcomes and sends still awaiting a confirmed Sent Items copy.

## Removed recovery

- Removed follow-ups are hidden from ordinary views.
- **Restore selected** returns them as Paused; it never resumes automatic sending.
- **Permanently purge** requires typing `PURGE` and deletes Organizer configuration, linked-message metadata, occurrences and history.
- Outlook emails are never deleted by restore or purge.

## Backup controls

- **Create verified backup** makes a manual snapshot in `D:\FollowupOrganizer\data\Backups`.
- **Verify latest backup** checks ZIP integrity, required recovery instructions and the SQLite database integrity.
- Backup verification reports the contained follow-up count.
- Restoring a database still requires Outlook and Organizer processes to be closed; V2.4 intentionally does not overwrite the live database from inside the running engine.

## Install over 2.3.7

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop Organizer background processes.
3. Replace `engine\core.py`, `engine\maintenance.py`, `engine\service.py` and `engine\version.py`.
4. Confirm `cFOListEvents` is already installed under **Class Modules**. If not, import `vba\cFOListEvents.cls` using **File -> Import File**.
5. Replace the complete existing **Organizer** UserForm code with `vba\Organizer-code.txt`.
6. Choose **Debug -> Compile VBAProject**, save, and restart Outlook.

No database migration is required. Existing data remains in `D:\FollowupOrganizer\data`.
