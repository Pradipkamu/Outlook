# Windows acceptance checklist

Record Outlook version, Office bitness, Python version and the installed pywin32 version before testing. Use Draft mode and test mailboxes you control. This checklist has not been executed in the build environment.

| Check | Expected result | Result |
| --- | --- | --- |
| Debug → Compile VBAProject | No compile errors in all imported modules and Organizer | Pending |
| Old organizer stopped | No old filing/case event handlers running | Pending |
| Service and VBA startup | One coordinator/watchdog; monitor heartbeat stays current for at least 5 minutes | Pending |
| Background task event / old reminder dismissal | No monitor reminder popup; background heartbeat remains current; unrelated reminders are retained | Pending |
| Received and sent enrollment | Correct mail, recipients and sending account; one follow-up per thread | Pending |
| Invalid dates / recipients | Invalid day/month combinations, past schedule, own recipient and display-name-only recipient rejected | Pending |
| Date and minute selectors | Every minute selectable; saved schedule appears in IST | Pending |
| Draft creation | One draft; correct sender, recipients, reply context and first-step text; no submitted increment | Pending |
| Sequence steps | Next occurrence uses next message; last message repeats; maximum applies | Pending |
| Human send / Sent Items | One submission, stored sent locator, future schedule; no duplicate draft | Pending |
| Incoming answer | Pause before review prompt; correct reply available in Conversation | Pending |
| Review later | Remains paused; review deadline and overdue prompt work | Pending |
| Close / Continue | Close persists; Continue requires future date/time; deadline unchanged | Pending |
| Before-send Yes/No/Cancel | Only Yes opens settings; Cancel retains draft; no schedule before sent-copy confirmation | Pending |
| Generated reminder send | Enrollment question is bypassed; a closed/paused reminder cannot be sent through organizer permission check | Pending |
| Follow-up Mail folder | Drag adds Setup required; no sending until enrollment | Pending |
| Rule folder / bulk sync | Registered rule destination reconciled; duplicate events do not repeat review notifications | Pending |
| Independent messages with same subject | Not automatically clubbed; explicit Link associates only the selected mail | Pending |
| Moving an original | Missing locator pauses safely; explicit Link repairs it | Pending |
| Excel export/import | IDs preserved, changed config validated, stale/formula rows rejected, no partial commit | Pending |
| Excel close/remove | Closed status disables; deleting row alone does not silently alter runtime state | Pending |
| Multiple selection | Pause/Close affects precisely selected valid rows; invalid batch changes none | Pending |
| Out-of-office | Conservatively paused; never auto-closed | Pending |
| Emergency stop | Stop button and separate .cmd latch stop; restart does not clear it | Pending |
| Worker stall | Watchdog detects stale heartbeat and blocks new dispatch without killing Outlook | Pending |
| Lost service / missing account / offline Outlook | Visible blocked/Needs action status; no blind retry | Pending |
| Restart with unfinished claim | Uncertain occurrence remains held until reviewed | Pending |
| Burst limit | Attempts above configured ceilings latch global stop | Pending |
| Outlook closed past an Auto due time, reopened inside catch-up window | After a complete online scan, exactly one reminder sends; next time is one full interval after catch-up | Pending |
| Several intervals missed while Outlook closed | Only one occurrence sends; no catch-up burst | Pending |
| Draft overdue or Auto overdue beyond catch-up window | Status becomes Needs action; no automatic send | Pending |
| Reply arrived while Outlook was closed | Reply scan changes record to Review before catch-up; no reminder sends | Pending |
| Backup | Backup archive opens; data snapshot exists and contains STOP | Pending |
| 100+ daily incoming mails | UI remains responsive; scan pass completes; no reply missed in controlled sample | Pending |

Only after these pass should you enable Auto globally and on a single test follow-up. Verify one automatic reminder end-to-end before enrolling operational recipients. A native Outlook security prompt may require local policy/IT configuration; do not bypass organizational controls.

## Automatic startup update checks

- Pending: Install-Automatic-Startup.cmd creates one current-user shortcut, including when the folder path contains spaces.
- Pending: Sign out/in or restart; both processes start with no console windows and no duplicate instances.
- Pending: Sign in before opening Outlook; worker waits without creating a false safety stop.
- Pending: Open Outlook before the service is ready; updated VBA monitor reconnects on its next tick.
- Pending: A pre-existing safety stop survives sign-in; removal deletes only this shortcut.

## Data location update checks

- Pending: Python, VBA, background startup and emergency stop all use D:\FollowupOrganizer\data.
- Pending: Migration refuses while old or new organizer/watchdog process locks are held.
- Pending: Migrated reminder counts/history/token and stop state match the old location; original data remains intact.

Background update checks: verify the Organizer form compiles, legacy FOClock reminders are disabled/dismissed, the non-reminding pulse runs for five minutes, and real replies still notify without a modal dialog in the background callback. All Windows checks remain pending.
