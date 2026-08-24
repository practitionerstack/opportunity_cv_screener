# Scale-Core Architecture Note

v2.7 remains a Streamlit MVP. It is designed to validate intake and scoring correctness before high-volume processing.

For thousands to millions of CVs, the next architecture should separate:

- intake/object storage
- queue-based batch processing
- document extraction workers
- OCR workers for scanned PDFs
- scoring service
- audit database
- duplicate index
- report generation workers
- client dashboard

The scoring rules validated in v2.7 should become the auditable scoring core before scaling infrastructure is added.
