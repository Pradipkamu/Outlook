# Classic Outlook Follow-up Organizer 2.5.5

Python background engine and Outlook VBA interface for scheduled email follow-ups, conversation review, reminders and Excel configuration exchange.

- [Installation and everyday use](START-HERE.md)
- [Upgrade from 2.2 and catch-up sending](UPDATE-2.3.md)
- [Windows verification checklist](WINDOWS-CHECKLIST.md)

Program folder: `D:\FollowupOrganizer`. Runtime database, settings, logs, exports and backups: `D:\FollowupOrganizer\data`. VBA form name: `Organizer`.

Requires Windows, Classic Outlook, an Outlook mail profile, Python with pywin32, and India Standard Time. Excel is required for workbook import/export. Outlook must remain open and the PC awake for monitoring.

Version 2.3 adds controlled catch-up sending when Outlook reopens. One overdue Auto occurrence can send within a configurable window (24 hours by default); missed intervals are never replayed as a burst. Draft mode, replies, safety stops, uncertain outcomes, limits and the global/per-follow-up Auto gates remain protected.

Version 2.3.1 displays Details / History timestamps in India Standard Time (`UTC+05:30`) while retaining UTC internally for safe scheduling and audit storage.

Version 2.3.2 applies the same IST display to the Conversation `Received` column, including older stored timestamps without an explicit offset. The Organizer window title now reads the running engine version instead of using a hard-coded caption.

Version 2.5.3 retains the Important Mail register and approved mailbox linking from V2.5.2, and adds automatic recovery when Classic Outlook closes or restarts. Known RPC disconnects now wait and reconnect without creating a nuisance safety latch; uncertain sends and genuine worker failures remain protected.

Version 2.5.4 preserves an existing follow-up's saved date and time when **Edit / Continue** is opened. The schedule changes only when the user edits those controls and saves.

Version 2.5.5 recovers linked emails whose Outlook location changed after a move, synchronization or restart by using their stable Internet Message ID, then refreshes the stored locator. It also removes the worker's dependency on temporary `gen_py` wrapper files.

Run automated tests from this folder:

```console
python -m unittest discover -s tests -v
```

Windows Outlook/VBA checks are required in addition to these Python tests. The source package contains no runtime mailbox data or credentials. Do not commit your runtime data, backups, tokens or virtual environment.
