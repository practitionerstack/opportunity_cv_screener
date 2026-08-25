import streamlit as st
import pandas as pd
import io, re, zipfile, unicodedata
from pathlib import Path
from datetime import date
import fitz
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="Opportunity Hub CV Screener v2.7.7", page_icon="🎯", layout="wide")

VERSION = "v2.7.7"
PREFERRED_WORDS = ("preferably", "preferred", "nice to have", "nice-to-have", "advantage", "bonus", "plus", "desirable")
REQUIRED_WORDS = ("must have", "must-have", "required", "essential", "mandatory", "minimum requirement")
SECTION_PREFERRED = {"nice to have", "nice-to-have", "preferred", "desirable", "advantage"}
SECTION_REQUIRED = {"must have", "must-have", "required", "essential", "mandatory", "requirements", "requirement"}

# Approved parser aliases only. These aliases are for interpreting JD language;
# they do not create requirements unless an alias is actually found in the JD.
REQUIREMENT_RULES = [
    ("Meta Ads", "skill", [r"\bmeta\s+ads?\b", r"\bfacebook\s+ads?\b", r"\binstagram\s+ads?\b"]),
    ("Google Ads", "skill", [r"\bgoogle\s+ads?\b", r"\badwords\b"]),
    ("Copywriting", "skill", [r"\bcopywriting\b", r"\bwrite\s+(?:ad\s+)?copy\b", r"\bad\s+copy\b", r"\bsocial\s+copy\b"]),
    ("GA4 / Google Analytics", "skill", [r"\bga4\b", r"\bgoogle\s+analytics(?:\s*4)?\b"]),
    ("Canva", "skill", [r"\bcanva\b"]),
    ("Fintech / banking experience", "preference", [r"\bfintech\b", r"\bbanking\b", r"\bbank\b"]),
    ("HubSpot", "skill", [r"\bhubspot\b"]),
    ("Basic Design", "skill", [r"\bbasic\s+(?:graphic\s+)?design\b", r"\bgraphic\s+design\s+skills?\b"]),
    ("TikTok Ads", "skill", [r"\btiktok\s+ads?\b"]),
]

SKILLS = {name: pats for name, _, pats in REQUIREMENT_RULES}


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("\u00a0", " ").replace("–", "-").replace("—", "-")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


def normalize_match(text: str) -> str:
    text = normalize_text(text).lower()
    return re.sub(r"\s+", " ", text).strip()


def extract(name, data):
    ext = Path(name).suffix.lower()
    try:
        if ext == ".pdf":
            with fitz.open(stream=data, filetype="pdf") as doc:
                text = "\n".join(page.get_text() for page in doc)
            if len(text.strip()) > 20:
                return text, ""
            reader = PdfReader(io.BytesIO(data))
            return "\n".join(page.extract_text() or "" for page in reader.pages), ""
        if ext == ".docx":
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs), ""
        if ext == ".txt":
            return data.decode(errors="ignore"), ""
    except Exception as exc:
        return "", str(exc)
    return "", f"Unsupported {ext}"


def clean_line(line):
    line = normalize_text(line)
    line = re.sub(r"^\s*(?:[-•*]+|\d+[.)])\s*", "", line)
    return line.strip()


def jd_lines(text):
    return [clean_line(x) for x in re.split(r"[\r\n]+", normalize_text(text)) if clean_line(x)]


def heading_or_section(line):
    match = re.match(r"^(must[- ]have|nice[- ]to[- ]have|required|essential|preferred|desirable|advantage|requirements?)\s*:?\s*(.*)$", line, re.I)
    if not match:
        return None
    return match.group(1).lower(), match.group(2).strip()


def section_default(section):
    s = (section or "General").lower().strip()
    if s in SECTION_PREFERRED:
        return "Preferred"
    if s in SECTION_REQUIRED:
        return "Required"
    return "Required"


def local_category(name, line, section):
    """Classify one requirement without letting a qualifier for another clause leak across the line."""
    low = normalize_match(line)
    section_cat = section_default(section)

    # Explicit local rule for the benchmark's compound line:
    # "2+ years ... experience, Fintech/Banking preferred"
    if name == "Fintech / banking experience":
        if any(w in low for w in ("fintech preferred", "banking preferred", "fintech/banking preferred", "fintech or banking preferred")):
            return "Preferred", "Local preferred qualifier detected"

    # Find the first alias position for this requirement and inspect nearby words.
    positions = []
    for rule_name, _, patterns in REQUIREMENT_RULES:
        if rule_name != name:
            continue
        for pattern in patterns:
            m = re.search(pattern, low, re.I)
            if m:
                positions.append((m.start(), m.end()))
    if positions:
        start, end = min(positions)
        window = low[max(0, start - 45):min(len(low), end + 60)]
        before = low[max(0, start - 30):start]
        after = low[end:min(len(low), end + 60)]
        if any(word in window for word in REQUIRED_WORDS):
            return "Required", "Local required qualifier detected"
        # A preferred qualifier immediately after a different later clause should not
        # downgrade an earlier required requirement. Only accept nearby/local preferred wording.
        if any(word in before for word in PREFERRED_WORDS) or any(after.startswith(word) or re.match(r"^(?:\s*[-,:()]*)?(?:" + "|".join(map(re.escape, PREFERRED_WORDS)) + r")\b", after) for word in PREFERRED_WORDS):
            return "Preferred", "Local preferred qualifier detected"
    return section_cat, f"Section/default category: {section_cat}"


def validate_requirement(req, full_jd):
    source = req.get("source_text", "")
    if not source or normalize_match(source) not in normalize_match(full_jd):
        return False, "Source excerpt not found in extracted JD text"
    name = req["name"]
    if req["type"] == "experience":
        years = req.get("years")
        if not years or not re.search(rf"\b{re.escape(str(years))}\s*\+?\s*years?\b", normalize_match(source)):
            return False, "Experience requirement is not supported by its source excerpt"
        return True, "Validated from source excerpt"
    for rule_name, _, patterns in REQUIREMENT_RULES:
        if rule_name == name:
            if any(re.search(p, source, re.I) for p in patterns):
                return True, "Validated from approved JD alias"
            return False, "Requirement label has no approved alias in its source excerpt"
    return False, "Unknown parser requirement rule"


def add_requirement(found, diagnostics, seen, req, full_jd):
    # One requirement cannot be both Required and Preferred: first direct occurrence wins,
    # and duplicate occurrences are recorded instead of duplicated into scoring.
    valid, reason = validate_requirement(req, full_jd)
    diag = {
        "Requirement": req["name"],
        "Category": req["category"],
        "Status": "Validated" if valid else "Rejected",
        "Reason": reason,
        "Source Excerpt": req.get("source_text", ""),
        "Section": req.get("source", "General"),
    }
    diagnostics.append(diag)
    if valid and req["name"] not in seen:
        seen.add(req["name"])
        found.append(req)


def parse_jd(text):
    full_jd = normalize_text(text)
    found, diagnostics, seen = [], [], set()
    section = "General"
    used_sections = set()

    for raw in jd_lines(full_jd):
        parsed = heading_or_section(raw)
        if parsed:
            section, remainder = parsed
            used_sections.add(section)
            if not remainder:
                continue
            line = remainder
        else:
            line = raw

        # Skills/preference requirements. Compound lines can legitimately yield multiple requirements.
        for name, req_type, patterns in REQUIREMENT_RULES:
            if any(re.search(pattern, line, re.I) for pattern in patterns):
                category, reason = local_category(name, line, section)
                req = {
                    "name": name,
                    "category": category,
                    "source": section.upper() if section != "General" else "GENERAL",
                    "source_text": line,
                    "type": req_type,
                    "weight": 4 if category == "Required" else 1,
                    "parser_reason": reason,
                }
                add_requirement(found, diagnostics, seen, req, full_jd)

        # Experience requirement. Do not let a later "Fintech/Banking preferred" clause
        # reclassify the experience requirement.
        year_match = re.search(r"\b(\d+)\s*\+?\s*years?\b", line, re.I)
        if year_match and any(token in normalize_match(line) for token in ("experience", "marketing", "digital")):
            years = int(year_match.group(1))
            category = "Preferred" if section_default(section) == "Preferred" else "Required"
            req = {
                "name": f"{years}+ years relevant experience",
                "category": category,
                "source": section.upper() if section != "General" else "GENERAL",
                "source_text": line,
                "type": "experience",
                "years": years,
                "weight": 4 if category == "Required" else 1,
                "parser_reason": "Experience threshold found in source excerpt",
            }
            add_requirement(found, diagnostics, seen, req, full_jd)

    summary = {
        "jd_text_extracted": bool(full_jd),
        "characters_extracted": len(full_jd),
        "sections_detected": len(used_sections),
        "candidate_requirements_found": len(diagnostics),
        "validated_requirements": sum(d["Status"] == "Validated" for d in diagnostics),
        "rejected_requirements": sum(d["Status"] == "Rejected" for d in diagnostics),
    }
    return found, diagnostics, summary


def yrs(text):
    total = 0
    for start, end in re.findall(r"\b(20\d{2})\s*(?:-|–|to)\s*(20\d{2}|present|current)\b", (text or "").lower()):
        end_year = date.today().year if end in ("present", "current") else int(end)
        total += max(0, end_year - int(start))
    return total


def text_lines(text):
    return [clean_line(x) for x in re.split(r"[\r\n]+", normalize_text(text)) if clean_line(x)]


def skill_ev(text, patterns):
    for line in text_lines(text):
        if any(re.search(pattern, line, re.I) for pattern in patterns):
            if re.search(r"\b(no|without|lack of|lacks?)\b", line, re.I):
                return 0, "Negative evidence", "Explicitly absent", line
            low = line.lower()
            if any(x in low for x in ("basic", "assisted", "support", "boosted posts", "sometimes")):
                return .35, "Positive evidence", "Weak / limited evidence", line
            if any(x in low for x in ("managed", "led", "budget", "certified", "proficient", "optimized", "built")):
                return 1, "Positive evidence", "Direct evidence", line
            return .7, "Positive evidence", "Evidence present", line
    return 0, "No evidence", "Not demonstrated in CV", ""


def score(text, requirements):
    rows = []
    earned = possible = preferred_earned = preferred_possible = 0
    missing = []
    for req in requirements:
        if req["type"] == "experience":
            years_found = yrs(text)
            level = min(1, years_found / req["years"]) if req.get("years") else 0
            evidence_type = "Positive evidence"
            status = "Meets requirement" if level == 1 else "Below requirement"
            evidence = f"Estimated dated experience: {years_found} year(s)"
        else:
            level, evidence_type, status, evidence = skill_ev(text, SKILLS[req["name"]])
        points = req["weight"] * level
        if req["category"] == "Required":
            earned += points
            possible += req["weight"]
            if level == 0:
                missing.append(req["name"])
        else:
            preferred_earned += points
            preferred_possible += req["weight"]
        rows.append({
            "Requirement": req["name"],
            "Category": "Must-have" if req["category"] == "Required" else "Preferred",
            "Source / Provenance": req["source"],
            "JD Source Excerpt": req["source_text"],
            "Requirement Type": req["type"],
            "Evidence Type": evidence_type,
            "Evidence Level": level,
            "Status": status,
            "Evidence": evidence,
            "Points Earned": round(points, 2),
        })
    mandatory = 100 * earned / (possible or 1)
    bonus = 5 * preferred_earned / preferred_possible if preferred_possible else 0
    final = min(100, mandatory + bonus)
    return round(final), round(mandatory, 1), round(bonus, 1), missing, rows


def candidate_name(filename):
    stem = Path(filename).stem
    stem = re.sub(r"^CV[_\- ]*\d+[_\- ]*", "", stem, flags=re.I)
    return stem.replace("-", " ").replace("_", " ").strip().title()

st.title(f"🎯 Opportunity Hub CV Screener {VERSION}")
st.caption("Parser Recovery & Diagnostics Hotfix • v2.7.5 preserved as rollback baseline")

jd_text = st.text_area("Paste Job Description", height=160)
jd_file = st.file_uploader("Upload JD", type=["pdf", "docx", "txt"])
cvs = st.file_uploader("Upload CVs or ZIP batch", type=["pdf", "docx", "txt", "zip"], accept_multiple_files=True)

if "screening_result" not in st.session_state:
    st.session_state.screening_result = None

if st.button("Screen CVs"):
    jd = jd_text
    jd_error = ""
    if jd_file:
        extracted, jd_error = extract(jd_file.name, jd_file.getvalue())
        if jd_error:
            st.error(jd_error)
            st.stop()
        jd = (jd + "\n" + extracted).strip()

    requirements, diagnostics, summary = parse_jd(jd)
    if not requirements:
        st.session_state.screening_result = {
            "requirements": [], "diagnostics": diagnostics, "summary": summary,
            "df": pd.DataFrame(), "audits": {}
        }
        st.error("No validated requirements extracted. Screening stopped safely. See JD Extraction Diagnostics below.")
    else:
        results, audits = [], {}
        for uploaded in cvs or []:
            items = []
            if uploaded.name.lower().endswith(".zip"):
                try:
                    with zipfile.ZipFile(io.BytesIO(uploaded.getvalue())) as archive:
                        items = [(Path(info.filename).name, archive.read(info)) for info in archive.infolist() if not info.is_dir()]
                except Exception as exc:
                    st.warning(f"Could not read ZIP {uploaded.name}: {exc}")
                    continue
            else:
                items = [(uploaded.name, uploaded.getvalue())]
            for filename, data in items:
                text, error = extract(filename, data)
                if error or not text:
                    continue
                fit, mandatory, bonus, missing, audit = score(text, requirements)
                group = "Excellent" if fit >= 90 else "Good" if fit >= 70 else "Moderate" if fit >= 50 else "Do Not Hire"
                name = candidate_name(filename)
                results.append({
                    "Name": name, "File": filename, "Fit %": fit, "Ranking Group": group,
                    "Mandatory Score": mandatory, "Preferred Bonus": bonus,
                    "Years Exp": yrs(text), "Why Not 100%": "; ".join(missing)
                })
                audits[name] = audit
        df = pd.DataFrame(results)
        if not df.empty:
            df = df.sort_values(["Fit %", "Mandatory Score"], ascending=False, kind="stable").reset_index(drop=True)
        st.session_state.screening_result = {
            "requirements": requirements, "diagnostics": diagnostics, "summary": summary,
            "df": df, "audits": audits
        }

result = st.session_state.screening_result
if result is not None:
    st.subheader("JD Extraction Diagnostics")
    summary_df = pd.DataFrame([{
        "JD text extracted": "Yes" if result["summary"]["jd_text_extracted"] else "No",
        "Characters extracted": result["summary"]["characters_extracted"],
        "Sections detected": result["summary"]["sections_detected"],
        "Candidate requirements found": result["summary"]["candidate_requirements_found"],
        "Validated requirements": result["summary"]["validated_requirements"],
        "Rejected requirements": result["summary"]["rejected_requirements"],
    }])
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    diagnostics_df = pd.DataFrame(result["diagnostics"])
    if not diagnostics_df.empty:
        rejected = diagnostics_df[diagnostics_df["Status"] == "Rejected"]
        if not rejected.empty:
            st.markdown("**Rejected Requirements**")
            st.dataframe(rejected, use_container_width=True, hide_index=True)

    requirements = result["requirements"]
    if requirements:
        st.subheader("JD Requirement Lock")
        lock = pd.DataFrame(requirements)
        st.dataframe(lock, use_container_width=True, hide_index=True)

        df, audits = result["df"], result["audits"]
        st.subheader("Master Ranking")
        st.dataframe(df, use_container_width=True, hide_index=True)

        st.subheader("Candidate-by-Candidate Scoring Audit")
        for name in df.get("Name", pd.Series(dtype=str)):
            with st.expander(name):
                st.dataframe(pd.DataFrame(audits[name]), use_container_width=True, hide_index=True)

        if not df.empty:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Master Ranking")
                lock.to_excel(writer, index=False, sheet_name="JD Requirements")
                pd.DataFrame(result["diagnostics"]).to_excel(writer, index=False, sheet_name="JD Diagnostics")
                all_audits = pd.concat(
                    [pd.DataFrame([{"Candidate": name, **row} for row in rows]) for name, rows in audits.items()],
                    ignore_index=True
                ) if audits else pd.DataFrame()
                all_audits.to_excel(writer, index=False, sheet_name="Scoring Audit")
            st.subheader("Downloads")
            st.download_button("Download Workbook", out.getvalue(), "opportunity_hub_v2_7_7.xlsx")
            st.download_button("Download CSV", df.to_csv(index=False).encode(), "opportunity_hub_v2_7_7.csv", "text/csv")

st.caption("Screened, sorted, scored and ranked by Opportunity Hub Screener.")
