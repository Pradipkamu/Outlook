# Outlook RPC reconnect correction for V2.5.2

Replace only `D:\FollowupOrganizer\engine\outlook_worker.py`, then close Outlook and restart the Organizer background processes.

The correction treats known Outlook RPC disconnects as normal lifecycle events. When Outlook closes, restarts or updates, the worker discards all stale COM proxies, reports **Waiting for Classic Outlook**, and reconnects without creating the global safety latch. No reminder is dispatched until the fresh Outlook session completes a full mailbox scan.

The safety latch intentionally remains for uncertain send outcomes, unfinished claims, loop/scan ceilings, database or unknown worker failures, prolonged before-send setup, and the manual Stop button.

An existing latch from an earlier RPC error must be cleared once with **Settings > Review and resume** after checking Outbox, Drafts and Sent Items. Future known RPC disconnects will not recreate it.
