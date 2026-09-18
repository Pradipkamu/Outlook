# Update to Follow-up Organizer 2.5.0

V2.5 adds an Important Mail register that is independent of reminder follow-ups.

## Important Mail

- Save a selected Inbox or Sent Items email without creating a reminder.
- Classify with 10 defaults: Quality, Production, Customer, Vendor, Purchase, Finance, Project, Management, Personal and Other.
- Type a new category when needed, up to 100 total categories.
- Add Normal, High or Critical priority, search tags, a private note and an optional review date.
- Search, open, edit, archive, permanently remove the bookmark, relink a moved email, or convert it into a follow-up.
- Outlook receives `FO Important` and `FO <Category>` category labels. The Organizer stores references and metadata; it does not copy message bodies or attachments.

## Update

Close Outlook and Organizer background processes, preserve the `data` folder, replace the V2.5 program files, import the updated VBA modules/form code, compile the VBA project, run `Setup.cmd`, then start Classic Outlook. The database table and default categories are created automatically.
