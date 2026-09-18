# Validation — release 2.3, 16 September 2026

67 Python tests passed with `python -m unittest discover -s tests -v`. Python source compilation passed. Tests cover the previous scheduler, reply matching, stop/loop guards, Excel transactions, fake COM dispatch, startup and migration behaviors plus:

- One overdue Auto occurrence is claimed inside the configured catch-up window and cannot be claimed twice.
- Catch-up schedules the next occurrence one complete interval after the actual catch-up claim, without replaying missed intervals.
- Draft-mode, disabled catch-up and Auto records beyond the window still require review.
- A reply scanned after Outlook reopens changes the record to Review before catch-up can claim it.
- Existing V2.2 settings automatically receive the safe 24-hour catch-up defaults.
- Catch-up window input is bounded from 1 to 168 hours.

- Snooze preserves deadlines, paused status and global stop; cannot bypass reply review or an outstanding draft/send; observes holidays, stop dates and existing later schedules.
- Daily snapshots contain a readable consistent database, exclude token.txt, run once per day, retain 30 daily archives, preserve manual backups and retain prior snapshots on publish failure.
- Diagnostics exclude private test subjects, message text, recipients, notes, raw errors/logs and credentials.
- Health gives safety/user stops precedence and identifies waiting for Outlook.
- Migration into the nested data folder preserves history and program files.
- Dashboard submitted count excludes reviewed-but-unconfirmed occurrences.

The user reported the previous build working correctly. That is prior-version feedback, not validation of these new changes. This environment cannot execute Windows Outlook/Excel COM, VBA compilation, actual sending, process locks or Startup shortcuts. Mock tests are not a substitute for those checks. Follow UPDATE-2.3.md for the short release check and WINDOWS-CHECKLIST.md for broader validation when needed.

No live messages were sent, no user's Outlook data was accessed, and no automatic-sending setting was enabled during development. The background monitor continues to use a non-reminding task and guarded VBA task-change events. No WinAPI timer or monitor reminder popup was introduced.

Version 2.3 retains the V2.2 unlink, address-book and appearance tests and adds catch-up scheduler coverage. The worker still performs a complete online scan before claiming due work, and dispatch retains its second stop/status check. Native Outlook scanning, VBA rendering and actual sending remain Windows checks.
