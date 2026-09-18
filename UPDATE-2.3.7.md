# Update to Follow-up Organizer 2.3.7

The lower dashboard area is now a **Selected follow-up summary** instead of a static information box.

- Refreshes automatically when a dashboard row is selected.
- Shows status, priority, deadline, next follow-up, sent count, recipients and next action.
- Shows an error prominently when present.
- Shows the latest three history events in newest-first IST order.
- Uses status-aware text colours: green for Active, orange for attention, grey for Paused/Closed and red for errors.
- Uses transparent background and no border, as requested.
- Conversation, Engine health, Diagnostics and full Details / History reuse the same clean panel.

## Install over 2.3.6

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop the Organizer background processes.
3. Replace `engine\version.py`.
4. Import the new class module `vba\cFOListEvents.cls` into the Outlook VBA project.
5. Replace the complete existing **Organizer** UserForm code with `vba\Organizer-code.txt`.
6. Choose **Debug -> Compile VBAProject**, save, and restart Outlook.

No database migration is required.
