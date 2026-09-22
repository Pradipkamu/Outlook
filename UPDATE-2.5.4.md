# Update to Follow-up Organizer 2.5.4

V2.5.4 corrects the **Edit / Continue** schedule behavior.

- Opening an existing follow-up now displays its saved next date and time in IST.
- Merely opening the editor no longer replaces the schedule with tomorrow at 09:30.
- The schedule changes only when the user changes the date or time and presses **Save**.
- New follow-ups continue to use the current clock time as their initial time default.

Replace `vba/Organizer-code.txt` in the `Organizer` form code module, then restart Outlook. Replace `engine/version.py` as well so the title reports V2.5.4.
