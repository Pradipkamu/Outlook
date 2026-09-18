# Update to Follow-up Organizer 2.3.1

This maintenance update changes the **Details / History** timestamp display from stored UTC (`+00:00`) to India Standard Time (`UTC+05:30`). The database continues storing timestamps in UTC; only the Outlook display is converted.

## Install over 2.3

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop the organizer background processes.
3. Copy the 2.3.1 program files into `D:\FollowupOrganizer`, preserving the existing `.venv` and complete `data` folder.
4. Run `Setup.cmd`. All 68 tests must pass.
5. Start Outlook and run `FollowupStart`.
6. Select a follow-up and open **Details / History**. Each event should appear like `16 Sep 2026 16:51:21 IST`, not `2026-09-16T11:21:21+00:00`.

No database migration and no VBA form-code replacement are required for this update.
