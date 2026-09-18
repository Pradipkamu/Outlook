# Update to Follow-up Organizer 2.3.2

This maintenance update displays the **Conversation → Received** column in India Standard Time (`UTC+05:30`). It also makes the Organizer window caption show the running engine version.

## Install over 2.3 or 2.3.1

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop the organizer background processes.
3. Copy the 2.3.2 program files into `D:\FollowupOrganizer`, preserving the existing `.venv` and complete `data` folder.
4. Run `Setup.cmd`. All 69 tests must pass.
5. In Outlook press **Alt+F11**, open the existing **Organizer** form, and replace its entire code with `vba\Organizer-code.txt`.
6. Choose **Debug → Compile VBAProject**, then save the VBA project.
7. Restart Outlook and run `FollowupStart`.

The Conversation heading should display **Received (IST)**, and dates should look like `16 Sep 2026 15:27:42 IST` instead of `2026-09-16T09:57:42+00:00`.

No database migration is required. UTC remains the internal audit and scheduling format.
