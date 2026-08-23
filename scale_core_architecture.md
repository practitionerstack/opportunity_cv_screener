# Enterprise Scale Core

Client Portal/API -> Object Storage -> Batch Manifest -> Queue -> Parallel Workers -> Results Database -> Ranking/Reporting

Required: per-file IDs, idempotency, resumable checkpoints, retry queue, dead-letter queue, audit trail, batch progress, horizontal worker scaling.

Do not process million-CV jobs inside a single Streamlit session.
