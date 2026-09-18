# Update to Follow-up Organizer 2.5.3

V2.5.3 is the stable V2.5.2 feature set plus a targeted Outlook RPC reconnect correction.

- Outlook close/restart/update no longer creates a safety latch for known RPC disconnect HRESULTs.
- All stale Outlook COM proxies are discarded before reconnecting.
- The worker reports **Waiting for Classic Outlook** until Outlook reopens.
- A complete mailbox scan is still required before reminders can be processed.
- Manual Stop, uncertain sends, unfinished claims, loop/scan ceilings, database failures and unknown worker errors still latch for review.

Replace `engine/outlook_worker.py`, restart the Organizer processes, check Drafts/Outbox/Sent Items, and use **Review and resume** once to clear a latch created by an older version.
