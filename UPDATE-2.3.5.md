# Update to Follow-up Organizer 2.3.5

After a new follow-up is saved from an existing Outlook email, Organizer now searches all configured Inbox folders and subfolders for related email.

- Matching ignores repeated `RE:`, `FW:` and `FWD:` prefixes, punctuation, case and repeated spaces.
- Every candidate requires an explicit **Yes**, **No** or **Cancel** decision.
- **Yes** links the message without changing the follow-up's reply source, schedule or status.
- **No** skips only that candidate.
- **Cancel** stops the remaining search; previously approved links remain.
- Existing linked messages are not proposed again.
- Safety ceilings: 200 folders, 20,000 messages and 100 candidate prompts per new follow-up.
- The search does not inspect sender email addresses, message bodies or attachments.

## Install over 2.3.4

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop the organizer background processes.
3. Replace the `engine` folder files with this release.
4. In Outlook press **Alt+F11** and replace the existing `FollowupOrganizer` standard-module code with `vba\FollowupOrganizer.bas`.
5. Replace the existing **Organizer** UserForm code with `vba\Organizer-code.txt`.
6. Choose **Debug -> Compile VBAProject**, save, and restart Outlook.

No database migration is required.
