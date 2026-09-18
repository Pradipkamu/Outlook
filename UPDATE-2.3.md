# Upgrade to Follow-up Organizer 2.3

Program: **D:\FollowupOrganizer**. Data: **D:\FollowupOrganizer\data**. Outlook form: **Organizer**.

## What changes

V2.3 can send one overdue **Auto** follow-up after Outlook reopens. The default catch-up window is 24 hours and can be changed from 1 to 168 hours in Settings.

Catch-up requires every normal sending gate:

- The record is Active and its own mode is Auto.
- Global automatic sending is enabled.
- STOP AUTOMATION is not latched.
- Outlook is online and a complete Inbox/Sent Items/monitored-folder scan has finished.
- No reply has changed the record to Review.
- No draft, unfinished claim or uncertain send is outstanding.
- Maximum reminders, stop date, work calendar and global send ceilings permit dispatch.

Only the next occurrence is claimed. Missed intervals are not replayed. After Sent Items confirms the catch-up message, the next reminder is scheduled one full interval after the catch-up claim time. Draft-mode records and Auto records older than the window become **Needs action**.

## Install over V2.2

1. Press **STOP AUTOMATION** and run **Backup-Organizer.cmd**.
2. Run **Remove-Automatic-Startup.cmd**, close Outlook, and restart Windows so no V2.2 Python process remains.
3. Extract the V2.3 ZIP into a temporary folder. Copy its contents into **D:\FollowupOrganizer**, replacing program files. Keep the existing `.venv` and `data` folders. Do not delete or overwrite `D:\FollowupOrganizer\data`.
4. Run **Setup.cmd**. All 67 tests must pass.
5. In Outlook press Alt+F11, open the existing **Organizer** form and replace its entire code with `vba\Organizer-code.txt`.
6. Choose **Debug → Compile VBAProject**, then save the VBA project.
7. Run **Install-Automatic-Startup.cmd** and run **FollowupStart** once from Outlook.
8. Open Settings. Confirm the new catch-up checkbox and `24` hour window. Keep Auto disabled until the controlled test below is complete.

No database migration is required. When the V2.3 engine first opens the existing database, it adds the new settings while preserving all previous preferences, follow-ups and history. Installation retains the existing STOP latch; it does not silently resume or enable global Auto sending.

## Controlled Windows test

Use a harmless conversation and recipient you control.

1. In Settings, temporarily use a working-hours window covering the test time. Enable global Auto and catch-up with a 24-hour window.
2. Create an Auto follow-up due three to five minutes ahead, then close Outlook before it becomes due.
3. Wait at least eleven minutes after its due time and reopen Outlook.
4. Confirm the engine first reports scanning, then Ready. Exactly one reminder should be requested and later confirmed in Sent Items.
5. Confirm Sent count becomes 1 and the next time is one complete configured interval after the catch-up time—not the series of missed times.
6. Send a reply while Outlook is closed for a second test record. On reopening, confirm the reply changes the record to Review and no catch-up reminder sends.
7. Test a Draft-mode overdue record; it must become Needs action and must not create a late draft automatically.
8. Set a one-hour catch-up window and test a record older than one hour; it must become Needs action.
9. Press STOP AUTOMATION and confirm catch-up cannot occur until explicit Review and resume.

After testing, restore your normal work calendar and catch-up window. Review Sent Items, Outbox, Drafts, dashboard history and `data\worker.log` before enabling this for live recipients.

## Existing V2.2 Needs action records

V2.3 does not automatically revive records that V2.2 already changed to **Needs action**. This avoids sending a mail that previously required review. For each such record, use **Edit / Continue**, choose a future time and save. Future missed schedules will use the V2.3 catch-up rule.

## Validation boundary

The 67 automated tests pass in the build environment, including reply-before-catch-up, duplicate-claim prevention and catch-up recurrence. Actual Classic Outlook COM, VBA compilation and real sending must be checked on your Windows computer. No real email was sent during development.
