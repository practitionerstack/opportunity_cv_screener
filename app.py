import io, re, json, zipfile
from datetime import datetime
from urllib.parse import urlparse, parse_qs

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="Opportunity Hub CV Screener v2.2", page_icon="🎯", layout="wide")

# ---------- text / file ingestion ----------
def normalize(s):
    return re.sub(r"\s+", " ", (s or "").lower()).strip()

def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((p.extract_text() or "") for p in reader.pages)

def extract_docx(data):
    d = Document(io.BytesIO(data))
    return "\n".join(p.text for p in d.paragraphs)

def extract_txt(data):
    return data.decode("utf-8", errors="ignore")

def extract_file_bytes(name, data):
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext == "pdf": return extract_pdf(data)
    if ext == "docx": return extract_docx(data)
    if ext in ("txt", "md"): return extract_txt(data)
    raise ValueError(f"Unsupported document type: .{ext}")

def google_drive_direct(url):
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if m: return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    qs = parse_qs(urlparse(url).query)
    if qs.get("id"): return f"https://drive.google.com/uc?export=download&id={qs['id'][0]}"
    return url

def filename_from_response(url, response):
    cd = response.headers.get("content-disposition", "")
    m = re.search(r'filename="?([^";]+)', cd, re.I)
    if m: return m.group(1)
    path = urlparse(url).path
    base = path.rsplit("/", 1)[-1] or "document.pdf"
    return base

def download_url(url):
    target = google_drive_direct(url.strip())
    r = requests.get(target, timeout=60, allow_redirects=True)
    r.raise_for_status()
    name = filename_from_response(target, r)
    ctype = r.headers.get("content-type", "").lower()
    # Google sometimes returns HTML login/permission page
    if "text/html" in ctype and len(r.content) < 2_000_000:
        raise ValueError("The link did not return the file. Make the file accessible to anyone with the link, then try again.")
    return name, r.content

# ---------- scoring ----------
SKILL_ALIASES = {
    "Meta Ads": [r"\bmeta ads?\b", r"\bfacebook ads?\b"],
    "Google Ads": [r"\bgoogle ads?\b", r"\badwords\b"],
    "GA4": [r"\bga4\b", r"\bgoogle analytics\b"],
    "Email Marketing": [r"\bemail marketing\b", r"\bemail campaigns?\b", r"\bnewsletters?\b"],
    "Canva": [r"\bcanva\b"],
    "Copywriting": [r"\bcopywriting\b", r"\bad copy\b", r"\bmarketing copy\b", r"\bsocial copy\b"],
    "HubSpot": [r"\bhubspot\b"],
    "Google Tag Manager": [r"\bgoogle tag manager\b", r"\bgtm\b"],
    "A/B Testing": [r"\ba/b testing\b", r"\bab testing\b", r"\bsplit testing\b"],
}
FINTECH = ["kuda", "opay", "carbon", "flutterwave", "paystack", "moniepoint", "paga", "gtbank", "access bank", "zenith bank", "firstbank", "sterling bank", "wema bank", "banking", "fintech"]

NEG_PATTERNS = [r"no\s+{term}", r"without\s+{term}", r"lack(?:s|ing)?\s+{term}", r"not\s+.*{term}"]
PARTIAL_WORDS = ["basic", "assisted", "support", "boosted posts", "exposure"]

def has_explicit_negative(text, label):
    term = re.escape(label.lower())
    return any(re.search(p.format(term=term), text) for p in NEG_PATTERNS)

def skill_evidence(text, skill):
    t = normalize(text)
    if has_explicit_negative(t, skill): return 0.0, "explicitly missing"
    found = any(re.search(p, t, re.I) for p in SKILL_ALIASES.get(skill, [re.escape(skill.lower())]))
    if not found: return 0.0, "not evidenced"
    if any(w in t for w in PARTIAL_WORDS):
        # only reduce if qualifier appears near the skill where practical
        for pat in SKILL_ALIASES.get(skill, []):
            m = re.search(r"(.{0,60})" + pat, t, re.I)
            if m and any(w in m.group(1) for w in PARTIAL_WORDS):
                return 0.5, "basic/limited evidence"
    return 1.0, "confirmed"

def extract_years(text):
    years = re.findall(r"\b(20\d{2})\s*[-–—to]+\s*(20\d{2}|present|current)\b", text, re.I)
    vals=[]
    now=datetime.now().year
    for a,b in years:
        end = now if b.lower() in ("present","current") else int(b)
        vals.append(max(0,end-int(a)))
    return max(vals) if vals else 0

def infer_name(text, fallback):
    lines=[x.strip() for x in text.splitlines() if x.strip()]
    for line in lines[:8]:
        if re.fullmatch(r"[A-Za-z][A-Za-z .'-]{2,60}", line) and not any(k in line.lower() for k in ["curriculum","resume","cv","experience","email","phone"]):
            return line.title()
    return fallback.rsplit('.',1)[0].replace('_',' ').replace('-',' ').title()

def parse_jd(jd):
    t=normalize(jd)
    role = "Screened Role"
    m=re.search(r"(?:role|position|job title)\s*[:#-]?\s*([^\n]{3,100})", jd, re.I)
    if m: role=m.group(1).strip()
    required=[]
    for s in SKILL_ALIASES:
        if any(re.search(p,t,re.I) for p in SKILL_ALIASES[s]): required.append(s)
    min_years=0
    ym=re.search(r"(\d+)\+?\s+years?", t)
    if ym: min_years=int(ym.group(1))
    location=""
    for city in ["lagos","abuja","port harcourt","ibadan","enugu","remote"]:
        if city in t: location=city.title(); break
    fintech = any(x in t for x in ["fintech","banking","financial services"])
    return {"role":role,"required":required,"min_years":min_years,"location":location,"fintech":fintech}

def score_candidate(name, cv_text, jd):
    spec=parse_jd(jd)
    t=normalize(cv_text)
    years=extract_years(cv_text)
    audit=[]; deductions=[]; score=100.0
    # Skills: only requirements explicitly found in JD
    for skill in spec["required"]:
        ev, reason=skill_evidence(cv_text, skill)
        if ev==1: audit.append(f"{skill}: confirmed")
        elif ev==0.5:
            score-=6; deductions.append(f"Limited {skill}: -6pts"); audit.append(f"{skill}: limited")
        else:
            score-=12; deductions.append(f"No {skill} evidence: -12pts"); audit.append(f"{skill}: missing")
    if spec["min_years"] and years < spec["min_years"]:
        score-=15; deductions.append(f"Less than {spec['min_years']} years: -15pts")
        audit.append(f"Experience: {years} years estimated")
    else: audit.append(f"Experience: {years} years estimated")
    if spec["fintech"]:
        if any(x in t for x in FINTECH): audit.append("Industry: fintech/banking evidenced")
        else: score-=10; deductions.append("No fintech/banking evidence: -10pts"); audit.append("Industry: not evidenced")
    # location from common city mentions, only if JD specifies a city and not remote
    loc="Remote" if "remote" in t else ""
    for city in ["lagos","abuja","port harcourt","ibadan","enugu"]:
        if city in t: loc=city.title(); break
    if spec["location"] and spec["location"].lower() not in ("remote",) and loc and loc.lower()!=spec["location"].lower():
        score-=5; deductions.append(f"Not {spec['location']}-based: -5pts")
    elif spec["location"] and not loc:
        audit.append("Location: not confidently extracted")
    score=max(0, round(score))
    if score>=90: group="Excellent"; verdict="Excellent match; shortlist immediately."
    elif score>=70: group="Good"; verdict="Strong candidate; shortlist with review."
    elif score>=50: group="Moderate"; verdict="Possible fit; gaps need careful review."
    elif score>=30: group="Maybe"; verdict="Weak fit; consider only if market is thin."
    else: group="Do Not Hire"; verdict="Poor match; do not prioritize."
    matched=[]
    for s in spec["required"]:
        ev,_=skill_evidence(cv_text,s)
        if ev>0: matched.append(s)
    green=[]
    if any(x in t for x in FINTECH): green.append("Direct fintech/banking experience")
    if years>=spec["min_years"]>0: green.append(f"{years}+ years experience")
    if matched: green.append("Matched: " + ", ".join(matched[:3]))
    red="; ".join(deductions[:2]) if deductions else "None material"
    why="; ".join(deductions[:3]) if deductions else "Meets evidenced requirements"
    return {
        "Name":name,"Fit %":score,"Group":group,"2-Line Verdict":verdict,
        "Why Not 100%":why,"Red Flag":red,"Green Flag":"; ".join(green) or "Evidence requires review",
        "Years Exp":years,"Skills Match":", ".join(matched) or "None evidenced",
        "Visa/Location":loc or "Review","Audit Trail":" | ".join(audit)
    }

def excel_bytes(df):
    out=io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name="Summary")
        for group in ["Excellent","Good","Moderate","Maybe","Do Not Hire"]:
            sub=df[df["Group"]==group]
            sub.to_excel(w,index=False,sheet_name=group[:31])
    return out.getvalue()

def read_docx_bytes(title, df):
    d=Document(); d.add_heading("Opportunity Hub Candidate Screening Summary",0)
    d.add_paragraph(f"Role: {title}")
    d.add_paragraph(f"Candidates screened: {len(df)}")
    d.add_paragraph(f"Generated: {datetime.now().strftime('%d %B %Y %H:%M')}")
    d.add_heading("Executive Summary", level=1)
    for g in ["Excellent","Good","Moderate","Maybe","Do Not Hire"]:
        d.add_paragraph(f"{g}: {int((df['Group']==g).sum())}")
    d.add_heading("Top Candidates", level=1)
    for _,r in df.head(10).iterrows(): d.add_paragraph(f"{r['Name']} — {r['Fit %']}% — {r['2-Line Verdict']}")
    b=io.BytesIO(); d.save(b); return b.getvalue()

# ---------- UI ----------
st.title("🎯 Opportunity Hub — CV Screening Service v2.2")
st.caption("Evidence-based candidate screening with a mobile-friendly URL import fallback.")

if "candidates" not in st.session_state: st.session_state.candidates=[]

with st.expander("How to use this version", expanded=True):
    st.write("If mobile file upload fails, use the **Public File Links** tab. For Google Drive, set each file to **Anyone with the link** and paste the links. The server downloads the files directly, bypassing the phone-to-Streamlit uploader.")

jd_text = st.text_area("Paste Job Description", height=180, placeholder="Paste the full JD here...")
input_mode=st.radio("Document intake", ["Public File Links (recommended if uploads fail)", "Phone Upload", "Paste CV Text"], horizontal=False)

jd_loaded=""
candidates=[]

if input_mode=="Public File Links (recommended if uploads fail)":
    jd_url=st.text_input("Public JD file link (PDF/DOCX/TXT)")
    cv_urls=st.text_area("Public CV file links — one link per line", height=180, placeholder="https://drive.google.com/file/d/.../view\nhttps://...")
    if jd_url:
        try:
            n,b=download_url(jd_url); jd_loaded=extract_file_bytes(n,b); st.success(f"JD link fetched: {n}")
        except Exception as e: st.error(f"JD link failed: {e}")
    for url in [x.strip() for x in cv_urls.splitlines() if x.strip()]:
        try:
            n,b=download_url(url); txt=extract_file_bytes(n,b); candidates.append((n,txt)); st.success(f"CV fetched: {n}")
        except Exception as e: st.error(f"CV link failed: {url[:50]}… — {e}")
elif input_mode=="Phone Upload":
    st.warning("For Android reliability, add one CV at a time. If this uploader shows AxiosError again, switch to Public File Links above.")
    jd_file=st.file_uploader("Upload one JD", type=["pdf","docx","txt"], key="jd_single")
    if jd_file:
        try: jd_loaded=extract_file_bytes(jd_file.name,jd_file.getvalue()); st.success("JD ready")
        except Exception as e: st.error(f"JD read failed: {e}")
    cv_file=st.file_uploader("Upload one CV", type=["pdf","docx","txt"], key="cv_single")
    if cv_file:
        try:
            txt=extract_file_bytes(cv_file.name,cv_file.getvalue())
            candidates.append((cv_file.name,txt)); st.success("CV ready for this run")
        except Exception as e: st.error(f"CV read failed: {e}")
else:
    pasted_name=st.text_input("Candidate name")
    pasted_cv=st.text_area("Paste CV text", height=250)
    if pasted_cv.strip(): candidates.append((pasted_name or "Candidate",pasted_cv))

if st.button("Screen Candidates", type="primary"):
    jd_final=jd_text.strip() or jd_loaded.strip()
    if not jd_final: st.error("Provide a Job Description by pasting it, uploading it, or adding a public link.")
    elif not candidates: st.error("Provide at least one CV by link, upload, or pasted text.")
    else:
        rows=[]
        prog=st.progress(0)
        for i,(fname,txt) in enumerate(candidates,1):
            name=infer_name(txt,fname)
            row=score_candidate(name,txt,jd_final); row["Resume Link"]=fname
            rows.append(row); prog.progress(i/len(candidates))
        df=pd.DataFrame(rows).sort_values(["Fit %","Name"],ascending=[False,True]).reset_index(drop=True)
        st.session_state.results=df; st.session_state.jd=jd_final

if "results" in st.session_state:
    df=st.session_state.results
    st.subheader("Screening Results")
    st.dataframe(df.drop(columns=["Group","Audit Trail"], errors="ignore"), use_container_width=True)
    st.download_button("Download Excel workbook", excel_bytes(df), "candidate_screening_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download CSV", df.to_csv(index=False).encode(), "candidate_screening_results.csv", "text/csv")
    role=parse_jd(st.session_state.jd)["role"]
    st.download_button("Download executive summary DOCX", read_docx_bytes(role,df), "candidate_screening_summary.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with st.expander("Show score audit trail"):
        st.dataframe(df[["Name","Fit %","Audit Trail"]], use_container_width=True)
