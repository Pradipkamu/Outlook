# Upgrade to Follow-up Organizer 2.2

Program: **D:\FollowupOrganizer**. Data: **D:\FollowupOrganizer\data**. Outlook form: **Organizer**.

## Install over your working copy

1. Press **STOP AUTOMATION**, run **Backup-Organizer.cmd**, and export your existing Outlook VBA components as a backup.
2. Run **Remove-Automatic-Startup.cmd**, close Outlook and restart Windows to stop the old background processes. Do not start the old organizer again.
3. Extract this ZIP to a temporary folder. Copy its organizer folder contents into **D:\FollowupOrganizer**, replacing program files. Keep your existing `.venv` and `data` folder. The ZIP contains no live database or token.
4. Run **Setup.cmd**. All 60 Python tests should pass. If your working data is already in `data`, no migration is required. For older root/LOCALAPPDATA installations, follow the migration steps in UPDATE-2.1.md before starting.
5. In Outlook press Alt+F11. Open the existing **Organizer** form and replace its entire code with **vba\Organizer-code.txt**. Do not import this text as a standard module or create a second form. If upgrading from 2.1, the other VBA modules have not changed; older versions should follow UPDATE-2.1.md to replace those components too.
6. Choose **Debug → Compile VBAProject**, save, then run **Install-Automatic-Startup.cmd** from D:\FollowupOrganizer and run **FollowupStart** once in Outlook.
7. Try the checks below in Draft mode. Inspect pending send outcomes and choose **Settings → Review and resume** when ready. This update does not enable Auto or clear a safety stop.

## New controls

**Unlink email:** select a follow-up → Conversation → select a message → Unlink email. A confirmation page shows the message and a list of replacement reply sources. Leave “No replacement” selected for ordinary linked messages. If removing the email used as the reply source, choose another linked original/reply. The only remaining email cannot be unlinked; use Close selected to stop that follow-up. Generated reminders remain in history. Resolve outstanding drafts/uncertain sends before changing links.

Unlink removes the organizer association, not the Outlook email. It records an exclusion so normal scanning does not add that same email again, including after a move when its Internet Message-ID is available. Existing Outlook tracking properties are retained; the organizer checks the exclusion before using them. Use **Link email** to explicitly include it again. The schedule and Review/Paused/Closed state are preserved. Other messages in that conversation, including future replies, can still be linked normally.

**Current time:** each time Add Follow-up opens, Hour and Minute use the current Windows clock (IST), replacing 09:30. The existing first-date default remains **tomorrow**; choose today or another date with the selectors. Working hours still adjust the saved schedule. Edit / Continue retains its existing tomorrow 09:30 proposal so overdue work does not resume immediately.

**Fonts / text colors:** Settings → Fonts / text colors. Choose separate fonts, sizes (8–10), and text colors for Dashboard and Add/Edit Follow-up. Click Save appearance. Choices persist in the organizer database under `data` and are included in database backups. Navigation/action buttons keep their original styling. The fixed layout uses a limited size range to retain space for its controls.

**Outlook Contacts:** in Add/Edit Follow-up, click Outlook Contacts beside Subject. The standard Outlook address-book dialog lets you choose Contacts or another available address list, search, and assign To/CC recipients. Existing typed addresses are included. OK copies resolved SMTP addresses into the form; Cancel leaves it unchanged. Unsupported contacts without an SMTP address show an error without changing either field. Save Follow-up to commit the recipients. This picker does not send an email or change your source message.

The picker uses Microsoft's documented [SelectNamesDialog](https://learn.microsoft.com/en-us/office/vba/api/outlook.selectnamesdialog) and [Recipients property](https://learn.microsoft.com/en-us/office/vba/api/outlook.selectnamesdialog.recipients). Address lists and access remain subject to your Outlook profile and organizational policy.

## Short Windows check

- Add a Draft-mode follow-up: verify current hour/minute and tomorrow's date. Open it again later and confirm the clock default refreshes.
- Open Outlook Contacts; choose a contact for To and another for CC. Test Cancel, then OK. Verify addresses, account and saved configuration.
- Link a harmless additional email, then unlink it from Conversation. Refresh and restart: it should stay excluded. Explicitly link it again and confirm it returns.
- Try unlinking the reply source: no replacement must fail safely; selecting another linked original/reply must succeed. Verify subsequent draft quotes the intended source. Do not send that draft unless you intend to.
- Change each page's font/text color, save and reopen Outlook. Check labels, input fields and list readability at your Windows display scaling.
- Confirm a reply still pauses the follow-up and STOP AUTOMATION still blocks sending.

60 automated Python tests passed in the build environment. Outlook VBA compilation, the contacts dialog, visual layout and actual Outlook behavior cannot be executed here and require these Windows checks. No real email was sent during development.
