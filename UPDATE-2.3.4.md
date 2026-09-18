# Update to Follow-up Organizer 2.3.4

This maintenance update makes chronological display consistent:

- **Conversation** messages are sorted by received date in descending order.
- **Details / History** events are sorted by event date in descending order.
- Both views continue to display dates in IST.
- Missing or damaged dates from legacy linked messages are placed at the bottom.
- **Open mail / draft** opens the newest linked message after the ordering change.

## Install over 2.3.3

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop the organizer background processes.
3. Replace the `engine` folder files with this release.
4. In Outlook press **Alt+F11** and open the existing **Organizer** form.
5. Replace its entire code with `vba\Organizer-code.txt` from this release.
6. Choose **Debug -> Compile VBAProject**, save, and restart Outlook.

No database migration is required.
