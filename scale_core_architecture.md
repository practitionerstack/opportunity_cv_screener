# Enterprise Scale Core

Production path:

Object Storage
→ Batch Manifest
→ Queue
→ Parallel Workers
→ Results Database
→ Ranking and Reporting

## Required scale capabilities
- Resumable batches
- Per-file status
- Idempotent processing
- Checkpoints
- Retry queue
- Dead-letter handling
- Audit events
- Horizontal worker scaling

The Streamlit application is the control panel and small-batch interface. Large jobs should process through external storage and workers without rewriting the screening engine.
