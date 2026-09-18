# Update to Follow-up Organizer 2.3.3

This maintenance update fixes the Outlook Contacts errors **Property is read-only** and **There must be at least one name or contact group** caused by incorrect dialog pre-population.

## Install over 2.3.2

1. Press **STOP AUTOMATION** and run `Backup-Organizer.cmd`.
2. In Outlook press **Alt+F11** and open the existing **Organizer** form.
3. Replace its entire code with `vba\Organizer-code.txt` from this release.
4. Choose **Debug → Compile VBAProject**, then save the project.
5. Replace `engine\version.py` so the Organizer caption reports 2.3.3.
6. Restart Outlook and the Organizer background engine.

When the Contacts window opens, select at least one person and place the person in **To** or **Cc** before pressing **OK**. Press **Cancel** when no recipient is required.

No database migration is required.
