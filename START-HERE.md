# Follow-up Organizer 2.3

**Existing installation: follow UPDATE-2.3.md first.** This release stores runtime data in **D:\FollowupOrganizer\data** and keeps program files in **D:\FollowupOrganizer**. Form name: **Organizer**.

Your previous build was reported working. This update adds controlled catch-up sending after Outlook was closed. The V2.2 unlink, current-time, appearance and Outlook Contacts functions remain included. Sending visibility, snooze, backups and safety controls remain available. The updated Python tests pass; compile the replacement VBA and check the new setting on Windows before resuming automated work.

## What you can do

- Select a received or sent email and add a follow-up without creating a case.
- Select year/month/day and hour/minute from dropdowns. Invalid dates are rejected.
- Choose Daily, Urgent or Weekly, or adjust interval and maximum reminders; monthly recurrence is supported.
- Set an optional task deadline and stop date independently.
- Edit a message sequence with Previous, Next, Add message and Remove message. The last message repeats until the limit/stop date.
- Generate an Outlook draft, or explicitly opt a follow-up into Auto mode after testing.
- Group linked original mail, drafts, sent reminders and replies in the Conversation view.
- Pause on matching replies, review the actual mail, then Close, Continue through the editor, or Review later.
- Use Today, Awaiting Reply, Review Replies, Closed, search, and bulk Pause/Close.
- Drag mail into Inbox / Follow-up Mail: it appears as Setup required until you click Edit / Continue and choose its settings.
- Export reminder configuration to Excel and import edits through VBA. Set status Closed to remove a reminder from sending without deleting history.
- Use Stop Automation, the external watchdog, limits, send claims and backups.

## Requirements

Windows, Classic Outlook desktop with a configured mail profile, Excel desktop for Excel features, and Python 3.11 or newer with the Windows Python launcher (`py`). Windows must use India Standard Time (UTC+05:30) for this pilot. No external date-picker control is needed, and no API keys are required.

The first setup downloads pywin32 from the normal Python package index. Follow your organization's macro and application policies; do not disable organizational security controls. Both normal 32-bit and 64-bit Office are intended, but neither was compiled/tested in this Linux environment.

## 1. Extract and prepare

1. Extract the entire ZIP to a permanent local folder, for example `D:\FollowupOrganizer`. Do not run inside the ZIP or place the live database on a network/OneDrive folder.
2. Run `Setup.cmd`. It creates a local Python environment and runs the included tests. Stop if a test fails.
3. Open Classic Outlook normally.
4. Run `Install-Automatic-Startup.cmd` once to enable background startup at Windows sign-in. See AUTOMATIC-STARTUP.md. Alternatively, run `Start-Organizer.cmd` manually and keep both processes running.

Data will be created under `D:\FollowupOrganizer\data`. The service listens only on authenticated loopback port 8765. Do not publish/forward that port. Keep `token.txt` private and do not include it when sharing diagnostic logs.

## 2. Replace the previous case organizer carefully

In Outlook press Alt+F11. Before changing anything, export the old modules/classes/form using File → Export File, and save the Outlook VBA project.

The previous package used `MailOrganizer`, `MOTests`, `cMOApplication`, `cMOButton`, `cMOInbox`, and `cMOSent`. After exporting them, remove those old organizer components from this VBA project so their errors/event handlers cannot interfere. Keep unrelated VBA modules. Remove old organizer startup calls and event forwarding from ThisOutlookSession; preserve unrelated handlers. Do not re-enable the old automatic filing code.

Use the form name **Organizer**. Export the existing UserForm12 first, change its (Name) to Organizer, and replace its code with `vba\Organizer-code.txt`. Existing installations should follow BACKGROUND-MONITOR-UPDATE.md to replace all related references together. The form builds its controls at runtime.

Import these five files using File → Import File:

1. `vba\FollowupOrganizer.bas`
2. `vba\FOExcel.bas`
3. `vba\cFOEvents.cls`
4. `vba\cFOButton.cls`
5. `vba\cFOFolder.cls`

There are five import files plus the Organizer form code replacement. For a fresh project, insert a blank UserForm and set its (Name) to Organizer. If that rename produces the previous Path/File access error, stop and report it; do not delete the entire Outlook VBA project.

In ThisOutlookSession, merge the `FollowupStart` call from `ThisOutlookSession-snippet.txt` into your existing Application_Startup procedure. Do not create a duplicate procedure. Do not add another ItemSend handler; cFOEvents owns the organizer's sending events.

Choose **Debug → Compile VBAProject**. If it highlights a line, take a screenshot and do not enable Auto mode. This is the first required Windows gate. Save the project.

## 3. Start the Outlook monitor and add buttons

Run macro **FollowupStart** using Alt+F8. It connects to the local service, creates Inbox / Follow-up Mail if needed, and attaches background task-change events for the VBA heartbeat. It disables the previous monitor reminder; the worker updates a non-reminding background task. A first-start stop remains until you explicitly resume.

Use Outlook File → Options → Quick Access Toolbar (or Customize Ribbon → New Group), choose commands from Macros, and add:

- `AddFollowup`
- `FollowupDashboard`
- `StopFollowupAutomation`
- Optionally `MonitorFollowupFolder`, `ExportFollowupsToExcel`, `ImportFollowupsFromExcel`

These buttons are added manually to the toolbar; the ZIP does not install an Outlook add-in or alter the Ribbon automatically.

Run **FollowupDashboard**. In Settings, add folders used by mail rules. Inbox and Sent Items of accessible stores are included; folders you add are monitored individually, not recursively. Press Review and resume only after the engine has completed its initial scan. If it says VBA monitor stale, run FollowupStart and validate the monitor task before proceeding.

## 4. First pilot — Draft mode only

Use a harmless test conversation between mailboxes you control. You, rather than this package, choose and send the test emails.

1. Select a received email and click AddFollowup. Verify To and Send from, leave Mode = Draft, and choose a time 3–5 minutes ahead within working hours. Save.
2. If outside working hours, adjust your test work calendar in Settings first. The engine shifts due times into the permitted working window.
3. Wait for Draft ready. Open Conversation, select the generated draft and click Open mail / draft. Check recipients, sender account, text and quoted thread. Creating a draft must not increment the submitted count.
4. Send the test draft yourself. Confirm one Submitted entry and one future occurrence. Do not interpret Submitted as delivery confirmation.
5. Reply from the counterpart mailbox. Verify Review state and that no further reminders can be sent. Open the reply, choose Review later, and confirm it stays paused.
6. Choose Edit / Continue, select a future time, and save. Then Close it and confirm it stays closed after restart.
7. Test a new normal outgoing email. Yes should open settings; Cancel keeps the email open; No sends without enrollment. Tracking activates only after the original appears in Sent Items.
8. Drag another email into Follow-up Mail and verify Setup required. It must not produce a reminder until you complete enrollment.
9. Use Stop-Automation.cmd while the Outlook dashboard is open. Verify the stop remains after restarting the service. Review and resume must be required.

Complete `WINDOWS-CHECKLIST.md` before checking “Enable automatic sending” in Settings. Auto sending also requires Mode = Auto on that particular follow-up. Both gates must be enabled deliberately.

In V2.3, Settings also contains **When Outlook reopens, send one overdue Auto follow-up** and a catch-up window from 1 to 168 hours. The default is enabled with 24 hours. The engine first completes its Outlook reply/Sent Items scan, then claims at most one occurrence. It never sends every missed interval. Draft-mode records and Auto records beyond the window remain **Needs action**.

## Everyday actions

**Review a reply:** Review Replies → select row → Conversation → select reply → Open mail. If satisfied, Close selected; if incomplete, Edit / Continue and choose a new reminder time. Opening/reading alone does not close anything.

**Postpone:** Select one Active or Paused follow-up and use Snooze 1 hour, Snooze tomorrow, or Next working day. Choices start from the current time and shift into permitted working hours. Tomorrow keeps the current clock time where possible; Next working day uses the next working day’s opening time. A choice earlier than the current schedule is rejected. Deadlines, stop dates and paused status remain unchanged. Reply review and unresolved sends must be handled first. Use Edit / Continue for a custom date/time.

**Combine related messages:** select the intended follow-up in the dashboard, select the additional email in Outlook, then Link email. This explicitly associates it and updates the anchor used for the next reply. A message already linked elsewhere is rejected. Subject similarity alone does not link messages.

**Outstanding draft or uncertain send:** Check Drafts, Outbox and Sent Items. The organizer never blindly resends an ambiguous submission. Remove unwanted unsent generated drafts manually, then use Resolve occurrence; this consumes the occurrence and pauses the follow-up. Edit to choose a future schedule. Do not use Resolve as a substitute for checking Outlook.

**Excel:** Export creates a new timestamped .xlsx under the data folder and opens Excel. Pause automation with Stop before importing. Keep id/version untouched. Edit the existing rows; source-linked enrollment happens in Outlook. Close the workbook, then Import. All rows validate before any are committed. Deleting an Excel row does not disable the follow-up; set its status to Closed instead. Imports never resume closed/review/paused work and never clear the global stop. There is no separate workbook that silently overwrites runtime state on startup.

**Backup:** A consistent daily snapshot is created while the service is running; the latest 30 daily snapshots are retained. Manual snapshots are never rotated. For an additional snapshot, run Backup-Organizer.cmd. It takes a consistent database snapshot into the data folder's Backups directory and includes a restore stop latch. It excludes the bridge credential. Stop and close the service before any restore; retain the current database separately and review all pending sends before resuming.

## Safety behavior and practical limits

One service and one watchdog can run at a time. The database journals each send occurrence before attempting it. A repeated request for the same occurrence cannot claim it twice. On an uncertain send, the record becomes Needs action. Initial global limits are 5 automated submission attempts per minute and 50 per hour. The watchdog writes a persistent stop if worker progress is missing for 45 seconds; it never kills Outlook.

Stop blocks future automated work. It cannot recall a submitted email or guarantee cancellation of a COM call already executing. Close also cannot remove mail already in Outbox; inspect it explicitly. An incoming reply in the narrow interval after preflight/Send may still cross a reminder already being submitted.

The VBA monitor uses background task-change events, with no scheduled monitor reminder and no WinAPI callback timer. An internal task may remain visible in Tasks, but its ReminderSet property is False. Do not delete it while the worker is active. Validate event delivery on Windows; the three-minute stale-heartbeat gate remains. See BACKGROUND-MONITOR-UPDATE.md.

This is a desktop system. Outlook must be open, the user must be signed in, and the PC must be awake. Immediate NewMailEx processing is supplemented by periodic folder reconciliation. No code can display an immediate reply alert while the PC is off. When Outlook returns, an eligible overdue Auto record can send once within the configured catch-up window. Missed occurrences are not replayed; Draft-mode and older overdue records require review.

## Practical boundaries

- No cloud/Windows service, new Outlook support, arbitrary timezone UI, or shared-mailbox/delegate certification. Per-user automatic Windows startup is included.
- Calendar date selection uses year/month/day dropdowns; the preview's browser calendar appearance is not a native VBA control.
- Email recipient suggestions cover standard SMTP and Exchange user entries. Check distribution lists, aliases and delegated accounts explicitly; unresolved recipient suggestions require an SMTP address.
- Automatic/out-of-office replies pause conservatively and never close work. Specialized delivery-failure reports are not yet classified automatically; review them in Outlook.
- Excel is a controlled configuration exchange, not a continuously live workbook. New source-linked follow-ups are enrolled from Outlook, not by creating unlinked spreadsheet rows.
- Working-day review prompts are included; a separate personal task planner, custom alert sounds, bulk snooze, and manual link undo UI remain later enhancements.
- Existing case records are not automatically converted. Enroll the emails you want to follow up; original mail is not deleted or bulk moved by installation.

## Troubleshooting

Service unavailable: check Engine health and `D:\FollowupOrganizer\data\service-startup.log` or `worker.log`. Background processes do not require an open command window. Do not send token.txt or the full database for troubleshooting.

Unresponsive Outlook: use the separate Stop-Automation.cmd first. The watchdog should also latch a stop. Avoid force-closing Outlook until you have considered unsaved drafts. After recovery, check pending send outcomes before resuming.

Port 8765 already used: close the duplicate organizer first; do not disable other software blindly. The service uses a process lock, so a second instance should exit.

Compile error: send the highlighted line and module name. This package intentionally avoids the previous `Release` method name and uses the Organizer form name.

Technical foundations: [Outlook ItemSend](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.itemsend), [Items.Restrict](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict), [Items.ItemAdd](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.itemadd), and [ConversationID](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.conversationid). Folder events can miss bulk additions, which is why the worker also reconciles changes.
