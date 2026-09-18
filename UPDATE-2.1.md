# Upgrade to Follow-up Organizer 2.1

Use this sequence for your existing working installation. Program folder: **D:\FollowupOrganizer**. Runtime data folder: **D:\FollowupOrganizer\data**. Form: **Organizer**.

## One-time upgrade

1. In the current organizer, press **STOP AUTOMATION**. Run your existing **Backup-Organizer.cmd** and keep that backup. Export the existing Outlook VBA modules and Organizer form as a backup.
2. Run **Remove-Automatic-Startup.cmd**, close Outlook, then restart Windows. Removing the shortcut alone does not stop background processes; the restart closes them. Do not launch the old organizer again.
3. Extract this ZIP to a temporary folder. Copy the contents of its organizer folder into **D:\FollowupOrganizer**, replacing program files. Keep your existing `.venv` and all existing data. Do not create an extra nested organizer folder. This ZIP contains no live database or token.
4. Run **Setup.cmd** from D:\FollowupOrganizer. The included tests must pass.
5. Run **Migrate-Existing-Data.cmd** before starting the service. It copies data from the old root folder (or the older LOCALAPPDATA location) into **data**, preserving the originals. If it reports a destination conflict, stop and share the message; do not delete either database. A fresh installation with no old database needs no migration.
6. Open Outlook's VBA editor (Alt+F11). After exporting backups, remove and reimport **FollowupOrganizer.bas**, **FOExcel.bas**, **cFOEvents.cls**, **cFOButton.cls**, and **cFOFolder.cls** from the new vba folder. Do not create duplicate modules. Open the existing form **Organizer**, replace its entire code with **Organizer-code.txt**, and retain its name. If it is still UserForm12, rename its (Name) to Organizer. Preserve unrelated VBA.
7. Keep one `FollowupStart` call in the existing `Application_Startup` procedure in ThisOutlookSession, using the included snippet as a reference. Choose **Debug → Compile VBAProject**, then save. Do not paste the form text into a standard module.
8. Run **Install-Automatic-Startup.cmd** from D:\FollowupOrganizer. Run **FollowupStart** once in Outlook, then open the dashboard. After future Windows sign-ins, the service and watchdog start in the background automatically.
9. Check the migrated follow-ups, recipients, schedules, history and any Drafts/Outbox items. The migration deliberately retains a safety stop. In Settings choose **Review and resume** only after checking them. Existing per-item Draft/Auto and global Auto settings are preserved.

## New controls

- **Sending columns:** next eligible schedule, last time a reminder was observed in Sent Items, actual submitted reminder count, and latest outcome. A consumed/reviewed uncertain occurrence is not counted as confirmed submitted. “In Sent Items” does not prove delivery or reading. Details / History shows failures and timestamps.
- **Snooze 1 hour:** postpone from now by an hour, shifted into working hours.
- **Snooze tomorrow:** same clock time tomorrow where possible, shifted into the working calendar.
- **Next working day:** beginning of the next permitted working day. Snooze applies to the current selected row, does not move an existing later schedule earlier, does not change deadlines, and keeps paused records paused. Review/uncertain-send states require review first.
- **Engine health:** current health, persistent stop reason, last successful online scan (UTC), version and latest daily backup. A safety stop takes precedence over a healthy scan.
- **Export diagnostics:** writes a JSON file under data\Diagnostics with version, health, counts and selected event types/times. It excludes email bodies, subjects, recipients, credentials and raw logs. Share that JSON for troubleshooting; do not share the database or token.
- **Daily backup:** service creates one snapshot per IST calendar day while running and retains the latest 30 daily archives in data\Backups. It retries a failed backup after an hour. Manual backups remain available and are never removed by rotation. There is no backup while the PC/service is off. Backups contain private follow-up data; they are not diagnostic exports.

## Short check before normal use

Check that the original follow-up count/history is present. In a Draft-mode test follow-up, try Snooze 1 hour and verify the deadline is unchanged. Verify a reply still pauses reminders. Open Engine health and confirm a successful scan and a daily backup filename after a minute. Export diagnostics and locate the file under data\Diagnostics. Check an existing sent reminder shows its count/history. Confirm no recurring monitor reminder popup reappears. Check automatic startup once at the next sign-in.

The Python suite passes in the build environment. The new VBA controls, Windows migration locks, backup permissions and Outlook behavior need this local check; no real emails were sent during development. You reported the previous version working correctly.
