# Opportunity Hub CV Screener v2.7.3

## v2.7.3 — Reliability Hotfix

This release fixes the v2.7.2 runtime crash caused by inconsistent requirement records reaching the scoring engine.

### Fixed
- One canonical requirement schema is now used across JD parsing, scoring, audit and exports.
- Pre-screen schema validation runs before any CV is scored.
- Malformed requirements stop screening safely with a readable message instead of crashing the Streamlit app.
- `requirement_rows()` now correctly returns and preserves both scoring requirements and parsed JD section metadata.
- The undefined `jd_sections` state bug is removed.
- Scoring uses validated requirement types only: `skill`, `preference`, `experience`, or `location`.
- Existing v2.7.2 controls remain: requirement lock, JD provenance, no hidden skill injection, explicit negative evidence, and distinct `Explicitly absent` / `Not demonstrated` / `Weak evidence` states.

### Regression test
Use the same Digital Marketing Executive JD and 10 CVs. Before trusting scores, inspect **JD Requirement Lock**. If the detected requirements are wrong, stop there; do not rely on the ranking.
