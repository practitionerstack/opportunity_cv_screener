# Opportunity Hub CV Screener v2.7

## v2.7 focus: Evidence Validation & Requirement Classification

This release fixes the two scoring problems exposed by v2.6:

1. **False-positive evidence:** `No GA4` or `No Google Ads` must not earn points merely because the skill name appears.
2. **Incorrect requirement classification:** requirements are sourced from the JD and tagged as **Must-have**, **Preferred**, or **General JD mention**.

### Core rules

- Must-have sections such as **Must Have**, **Required**, **Essential**, and **Minimum Qualifications** drive the primary score.
- Preferred sections such as **Nice to Have**, **Preferred**, **Bonus**, and **Advantage** add only a limited bonus.
- General JD mentions are never silently promoted to mandatory requirements.
- Every audit row records **Source / Provenance** so the client can see where the requirement came from.
- Explicit negative statements are recorded as **Explicitly absent / Negative evidence**.
- Strong positive evidence elsewhere in the CV can still count if a candidate has both a negative self-summary and genuine demonstrated experience.
- Duplicate files and duplicate extracted content are excluded from scoring.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Recommended deployment

Deploy `app.py` on Streamlit Community Cloud from GitHub.
