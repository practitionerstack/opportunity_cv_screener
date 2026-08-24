# Opportunity Hub CV Screener v2.6

## Reliability & Auditability Fix

This release focuses on:
1. Robust PDF intake using PyMuPDF first and pypdf as fallback.
2. Evidence-based scoring rather than loose keyword accumulation.
3. Candidate-by-candidate requirement, evidence, status and points audit trail.
4. Duplicate file/content detection.
5. Persistent Streamlit results after downloads.

## Supported intake
- JD: PDF, DOCX, TXT, CSV, XLSX/XLS or pasted text.
- CVs: PDF, DOCX, TXT, CSV, XLSX/XLS, individual uploads or ZIP batches.

## PDF behavior
Text-based PDFs are extracted through two parser paths. If both fail, or the PDF is scanned/image-only, the application reports that clearly instead of silently treating the file as a valid CV. OCR is a later scale-layer addition.

## Scoring philosophy
The screener extracts explicit JD requirements, separates required from preferred requirements, scores direct evidence more highly than weak evidence, checks experience, records evidence used, and prevents a perfect score when required requirements have no evidence.
