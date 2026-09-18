**Release 2.1: follow UPDATE-2.1.md first for the data migration and complete replacement sequence.**

# Background monitor and Organizer form update

The recurring “Follow-up Organizer monitor” reminder is replaced with a background event. There is no one-minute monitor popup and no WinAPI callback timer. Real reply alerts and morning/end-of-day reviews remain enabled.

## Install this update

1. Close Outlook. Extract the updated package over `D:\FollowupOrganizer` while keeping `.venv` and runtime data. Restart Windows once so automatic startup loads the updated Python worker.
2. Open Outlook and press Alt+F11. Export the current form and the modules you are about to replace as backups.
3. Select the existing form `UserForm12`. In its Properties window, change **(Name)** to **Organizer**. Replace its code with `vba\Organizer-code.txt`. If you already named the form Organizer, keep that name. Do not create two copies. If VBA reports a path/file-access error while renaming, report it rather than deleting the VBA project.
4. Remove the old **FollowupOrganizer** module and **cFOEvents** class after exporting them. Import the updated `vba\FollowupOrganizer.bas` and `vba\cFOEvents.cls`. Keep FOExcel, cFOButton and cFOFolder; their interfaces are unchanged.
5. Choose **Debug → Compile VBAProject**, then save. Keep the existing `Application_Startup` call to `FollowupStart`.
6. Run **FollowupStart** once. It switches off reminders on old monitor tasks identified by the FOClock marker and dismisses only those active reminders. Other Outlook reminders are not suppressed. On future Outlook launches, normal startup performs this automatically.
7. Observe for at least five minutes: no “Follow-up Organizer monitor” reminder should appear, the VBA heartbeat should stay current, and a controlled test reply should still enter Review. Review existing safety stops before resuming automated sending.

## Background behavior

The independent worker updates a task named **Follow-up Organizer background monitor** approximately every 30 seconds while Outlook is connected. That task always has ReminderSet=False. It can remain visible in Tasks, but does not schedule a reminder or display itself.

The task change triggers a guarded VBA Items.ItemChange/ItemAdd handler, which updates the heartbeat and checks follow-up status. It does not write back to the task, preventing a self-triggering update loop. Checks initiated by this background event use a modeless dashboard for actionable reviews rather than a blocking message box. The independent watchdog and the three-minute VBA heartbeat gate remain in place.

No popup is shown merely because the background pulse occurs. A real incoming reply or a scheduled daily review can still bring the organizer to your attention, as requested. If the worker cannot connect to Outlook, changing the monitor mechanism does not fix that separate connection issue; inspect heartbeat.json / worker.log.

## Validation

40 automated tests pass locally, including two new tests for a non-reminding, rate-limited pulse and reuse of an existing tagged monitor task. Actual Outlook event delivery, legacy reminder dismissal and VBA compilation still require validation on Windows.

Implementation reference: [Microsoft Outlook Items.ItemChange event](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.itemchange).
