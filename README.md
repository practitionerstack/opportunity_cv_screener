# Opportunity Hub CV Screener v2.7.7

## Purpose
Parser Recovery & Diagnostics Hotfix.

This release changes only JD parsing, validation and parser diagnostics. It does not intentionally redesign the commercial workflow or ranking presentation.

## v2.7.7 fixes
- Separates JD normalization from source validation.
- Validates the source excerpt rather than requiring the display label to appear literally.
- Allows compound JD lines to yield multiple requirements from one real source excerpt.
- Supports approved aliases without internal skill injection.
- Applies qualifiers locally so a later `preferred` clause does not downgrade an earlier requirement.
- Keeps one requirement in one category only: Required or Preferred.
- Shows JD Extraction Diagnostics and rejected requirements with reasons.
- Stops safely when no requirements validate.

## Benchmark expectation
The existing benchmark JD should recover these ten requirements:
1. Meta Ads — Required
2. Google Ads — Required
3. Copywriting — Required
4. GA4 / Google Analytics — Required
5. Canva — Required
6. 2+ years relevant experience — Required
7. Fintech / banking experience — Preferred
8. HubSpot — Preferred
9. Basic Design — Preferred
10. TikTok Ads — Preferred

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
