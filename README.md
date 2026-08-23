# Opportunity Hub CV Screening Service v2.5

## v2.5 fixes
- JD upload: PDF, DOCX, TXT, CSV, XLSX/XLS
- Individual multi-CV upload
- ZIP CV upload
- Public links
- Evidence-only requirement scoring
- Requirement-by-requirement evidence matrix
- Unsupported 100% scores blocked
- Duplicate detection by exact file hash and document-content fingerprint
- Persistent results after downloads
- Master workbook plus Excellent, Good, Moderate and Do Not Hire sheets
- CSV, professional DOCX and PDF reports

## Run
pip install -r requirements.txt
streamlit run app.py

## Scale path
v2.5 is the single-session foundation. For thousands/millions of CVs, use object storage, batch manifests, asynchronous queues, autoscaling workers, a database-backed duplicate index, retries and a human-review queue.
