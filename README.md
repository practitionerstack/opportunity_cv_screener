# Opportunity Hub CV Screener v2.7.2

## v2.7.2 — Requirement Lock & Scoring Calibration

This release fixes the final scoring-core issues found in the controlled 10-CV regression test.

### What is fixed
- **Requirement lock:** only requirements explicitly extracted from the supplied JD's Must Have / Required or Nice to Have / Preferred sections can affect scoring.
- **No hidden requirement injection:** a recogniser dictionary cannot introduce TikTok Ads, SEO, A/B testing, or any other requirement unless that wording is actually present in the relevant JD section.
- **Requirement provenance:** every scored requirement shows its exact JD section and source excerpt.
- **Fair evidence states:** `Explicitly absent`, `Not demonstrated in CV`, and `Weak / limited evidence` are distinct.
- **Scoring calibration:** must-haves drive the main score; preferred items provide only a limited bonus.

### Controlled test expectation
Use the same JD and 10 CVs. First inspect **JD Requirement Lock**. If a requirement is not shown there, it cannot affect any candidate score.
