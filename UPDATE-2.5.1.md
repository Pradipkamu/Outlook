# Update to Follow-up Organizer 2.5.1

When a new follow-up is saved from an existing Outlook email, Organizer now searches the Inbox and Sent Items folder trees in every configured Outlook store. It identifies candidates using the Outlook Conversation ID first and the normalized subject as a fallback.

Each candidate is shown separately with Yes, No and Cancel choices. Nothing is linked without approval. Existing safeguards remain: generated reminder messages are excluded, already-linked messages are skipped, and folder/message/candidate ceilings stop runaway scanning.

The same approved search now runs after saving a new Important Mail bookmark. Related bookmark emails are grouped in the bookmark record and marked with the same Outlook Important Mail category; they do not become follow-ups.
