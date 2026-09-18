# Automatic startup update — 14 September 2026

**Data-folder update:** existing installations must follow DATA-LOCATION-UPDATE.md before restarting. The runtime data location is now D:\FollowupOrganizer\data. For release 2.1 use UPDATE-2.1.md before these background-startup instructions.

## One-time installation

1. Extract this updated package over your existing permanent organizer folder. Keep its existing `.venv` directory. If this is a new installation, run Setup.cmd first.
2. In Outlook VBA, export the existing FollowupOrganizer module as a backup, remove that module, and import the updated `vba\FollowupOrganizer.bas`. Compile and save. For the latest background-monitor update, also replace cFOEvents and rename the form as instructed in BACKGROUND-MONITOR-UPDATE.md. Keep the existing Application_Startup call to FollowupStart.
3. Run **Install-Automatic-Startup.cmd** once as your normal Windows user. It creates one Startup shortcut for that user and requests background launch immediately. No administrator rights are requested.
4. Restart Windows once to replace any previously running copies and verify the update. Sign in, open Classic Outlook normally, and check the follow-up dashboard. You should not need to run the organizer or watchdog CMD files again on subsequent sign-ins.

The installer creates `Follow-up Organizer.lnk` in your current user's Windows Startup folder. It points to the Python environment in the extracted package. Keep that folder in its permanent location; if you move it, rerun the installer from the new location. Repeating installation updates the same shortcut. Existing process locks prevent duplicate organizers/watchdogs.

## How it works

Both background processes launch without console windows. Outlook is not opened automatically; the organizer waits for you to open Classic Outlook. VBA startup attaches task-change events before connecting to the service. The worker triggers a background check after it connects to Outlook.

Automatic startup does not enable Auto mode or clear an emergency stop. If your dashboard was stopped before restarting, review the reason and use Review and resume. If Outlook closes unexpectedly during active work or a send outcome is uncertain, normal safety rules still apply.

The watchdog allows 45 seconds for its first startup heartbeat. After this startup grace period, the existing stale-heartbeat protection applies. Waiting for Outlook is a healthy waiting state, with sending unavailable. It is not treated as a COM processing failure.

Your PC must still be awake and signed in, and Outlook must be open for follow-up processing. This is a per-user desktop startup entry, not a Windows service or cloud scheduler.

## Remove automatic startup

Run **Remove-Automatic-Startup.cmd**. It removes only this organizer's shortcut; it does not delete reminders or stop currently running processes. Use Stop-Automation.cmd to block processing immediately. You can also open Windows Run, enter `shell:startup`, and delete only `Follow-up Organizer.lnk`.

## Troubleshooting

- If nothing starts, check `D:\FollowupOrganizer\data\startup-error.log`, `service-startup.log` and `watchdog-startup.log`.
- Processing errors continue to appear in `worker.log`; the dashboard displays safety stops.
- If Startup apps are disabled by Windows or organizational policy, enable this entry through the permitted Windows settings or ask your IT team. The installer does not override policies.
- If the VBA monitor is stale after sign-in, confirm that the updated FollowupOrganizer.bas was imported and Application_Startup still calls FollowupStart.
- Manual Start-Organizer.cmd remains available for diagnostics; duplicate processes exit using their existing locks.

Validation: five automated startup tests cover repeatable shortcut creation with paths containing spaces, missing environment handling, removal scope, background launch with stop preservation, and absent-Outlook versus other COM errors. These use mocks; actual Windows shortcut creation, sign-in launch and VBA startup timing require local validation.

Reference: [Microsoft — configure startup applications](https://support.microsoft.com/en-us/windows/experience/startup-boot/configure-startup-applications-in-windows) describes per-user Startup-folder shortcuts and their removal.
