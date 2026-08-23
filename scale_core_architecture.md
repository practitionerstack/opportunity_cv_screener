# Scale Architecture
Browser/API -> resumable uploads -> object storage -> queue -> extraction workers -> duplicate fingerprint index -> evidence scoring -> database/search -> reporting.

Controls: idempotent job IDs, per-file status, retries, dead-letter queue, audit evidence, duplicate hashes/content fingerprints, human review for borderline cases.
