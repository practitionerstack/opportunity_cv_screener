# Scale path

v2.6 is for immediate use. Millions of CVs should not be processed inside one Streamlit process.

Upgrade path:
1. Object storage for files.
2. Upload manifest instead of one HTTP request containing all CVs.
3. Queue-based ingestion.
4. Stateless extraction workers.
5. OCR worker pool for scanned PDFs.
6. PostgreSQL metadata and audit tables.
7. Hash-based duplicate detection plus similarity fingerprints.
8. Batch IDs, retries, checkpoints and dead-letter queues.
9. Separate scoring workers and report-generation workers.
10. Human-review queue for ambiguous evidence.
