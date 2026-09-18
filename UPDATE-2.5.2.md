# Update to Follow-up Organizer 2.5.2

This corrective update guarantees that an Outlook category or follow-up property is written only after the Organizer engine confirms that the candidate was successfully linked.

- **Yes + confirmed link:** store the link and apply its Organizer category/property.
- **No:** skip without changing the email.
- **Cancel:** stop without changing the current email.
- **Duplicate or rejected link:** count it as skipped and do not change the email.

Categories incorrectly applied by V2.5.1 are not removed automatically because Organizer cannot safely determine whether the user later assigned the same category intentionally.
