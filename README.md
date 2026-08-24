# Opportunity Hub CV Screener v2.7.4 — True JD Lock Hotfix

This release is a pure Job Description parsing and requirement-lock hotfix.

## v2.7.4 guarantees

- The supplied JD is the sole source of scoring requirements.
- Internal skill recognisers cannot inject requirements by themselves.
- TikTok Ads is extracted only when the exact phrase `TikTok Ads` appears in the JD source excerpt.
- `Preferably`, `Preferred`, `Nice to have`, `Advantage`, `Desirable`, `Bonus`, and `Plus` classify the specific requirement as Preferred.
- `Must have`, `Required`, and `Essential` classify the specific requirement as Required.
- A requirement has one final category only; it cannot be both Required and Preferred.
- Every requirement must retain an exact JD source excerpt containing the requirement or a recognised synonym.
- If provenance validation fails, screening stops safely before candidate scoring.

## Regression test

Run the same Digital Marketing Executive JD and 10 CV benchmark used for v2.7.1–v2.7.3.

Expected checks:

1. TikTok Ads must not appear unless those exact words are in the supplied JD.
2. `preferably fintech/banking` must be Preferred, not Required.
3. Zainab's `No GA4` and `No Google Ads` must remain Explicitly absent.
4. The JD Requirements sheet is the scoring contract.
