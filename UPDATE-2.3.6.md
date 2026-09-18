# Update to Follow-up Organizer 2.3.6

## Remove complete follow-up

- The dashboard now includes **Remove follow-up**.
- One or multiple selected follow-ups can be removed after explicit confirmation.
- Removed follow-ups disappear from all Organizer lists and cannot create further reminders.
- Replies and late worker errors cannot reopen a removed follow-up.
- Outlook emails are not deleted.
- Database history is retained for audit and recovery safety.
- Existing Draft or Outbox items must still be checked manually because submitted mail cannot be recalled.

## Preserve inline images in reminders

- Reminder text is now HTML-escaped and prepended using Outlook `HTMLBody`.
- The original reply HTML and CID references remain intact.
- Signature logos and other inline pictures should remain inline instead of appearing as `~WRD0000.jpg` attachments.
- No attachment is deleted by filename, so genuine user attachments are not accidentally removed.

## Install over 2.3.5

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. Close Outlook and stop the organizer background processes.
3. Replace the `engine` folder files with this release.
4. In Outlook press **Alt+F11** and replace the existing **Organizer** UserForm code with `vba\Organizer-code.txt`.
5. Choose **Debug -> Compile VBAProject**, save, and restart Outlook.

No database migration is required. Existing organizer data remains in `D:\FollowupOrganizer\data`.
