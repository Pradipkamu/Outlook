# Update to Follow-up Organizer 2.5.5

V2.5.5 corrects two Outlook restart and synchronization problems.

## Moved or resynchronized linked email

If Outlook can no longer open a saved `EntryID`, Organizer searches mailbox folders using the email's stable Internet Message ID. When found, it opens the message and refreshes the stored Outlook location automatically. The search is limited to 300 folders. If no safe exact match exists, Organizer asks the user to select the current message and use **Link email**; it never opens a subject-only match.

## Temporary gen_py cache

The worker now uses Outlook's active COM object directly. It no longer calls `gencache.EnsureDispatch`, so Windows cleanup of `%LOCALAPPDATA%\Temp\gen_py` cannot break mail access while Organizer is running.

Replace these files and restart Outlook and Organizer:

- `engine/outlook_worker.py`
- `engine/core.py`
- `engine/service.py`
- `engine/version.py`
- the `Organizer` form code using `vba/Organizer-code.txt`
