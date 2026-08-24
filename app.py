import streamlit as st
import pandas as pd
import hashlib, io, re, zipfile, tempfile
from datetime import date
from pathlib import Path
import fitz
from pypdf import PdfReader
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt

st.set_page_config(page_title="Opportunity Hub CV Screener v2.7", page_icon="🎯", layout="wide")

# ------------------------------ FILE INTAKE ------------------------------
def sha256(data): return hashlib.sha256(data).hexdigest()

def pdf_text(data):
    errors=[]
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            text="\n".join(p.get_text("text") for p in doc)
        if len(text.strip()) >= 20: return text.strip(), "PyMuPDF", None
        errors.append("PyMuPDF returned little or no text")
    except Exception as e: errors.append(f"PyMuPDF: {str(e)[:120]}")
    try:
        reader=PdfReader(io.BytesIO(data))
        text="\n".join((p.extract_text() or "") for p in reader.pages)
        if len(text.strip()) >= 20: return text.strip(), "pypdf", None
        errors.append("pypdf returned little or no text")
    except Exception as e: errors.append(f"pypdf: {str(e)[:120]}")
    return "", None, "PDF could not be read as text. It may be scanned/image-only and require OCR. " + " | ".join(errors)

def extract(name,data):
    ext=Path(name).suffix.lower()
    try:
        if ext==".pdf": return pdf_text(data)
        if ext==".docx":
            doc=Document(io.BytesIO(data)); parts=[p.text for p in doc.paragraphs]
            for t in doc.tables:
                for row in t.rows: parts.append(" | ".join(c.text for c in row.cells))
            return "\n".join(parts).strip(), "python-docx", None
        if ext in [".txt", ".md"]: return data.decode("utf-8",errors="ignore"), "text", None
        if ext in [".csv", ".xlsx", ".xls"]:
            bio=io.BytesIO(data); df=pd.read_csv(bio) if ext==".csv" else pd.read_excel(bio)
            return df.astype(str).fillna("").to_csv(index=False), "pandas", None
        return "", None, f"Unsupported file type: {ext or 'unknown'}"
    except Exception as e:
        return "", None, f"Could not extract text: {str(e)[:160]}"

def unpack(name,data):
    if Path(name).suffix.lower()!=".zip": return [(name,data)]
    out=[]
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for i in z.infolist():
                if not i.is_dir() and Path(i.filename).suffix.lower() in [".pdf",".docx",".txt",".csv",".xlsx",".xls"]:
                    out.append((Path(i.filename).name,z.read(i)))
    except Exception as e:
        return [(name,b"")]
    return out

# --------------------------- REQUIREMENT MODEL ---------------------------
SKILLS={
"Meta Ads":[r"\bmeta ads?\b",r"\bfacebook ads?\b",r"\binstagram ads?\b"],
"Google Ads":[r"\bgoogle ads?\b",r"\badwords\b"],
"Email Marketing":[r"\bemail campaigns?\b",r"\bemail marketing\b",r"\bnewsletter[s]?\b"],
"Copywriting":[r"\bcopywriting\b",r"\bwrite copy\b",r"\bad copy\b",r"\bsocial copy\b"],
"GA4 / Google Analytics":[r"\bga4\b",r"\bgoogle analytics\b"],
"Canva":[r"\bcanva\b"],
"HubSpot":[r"\bhubspot\b"],
"Basic Design":[r"\bbasic design\b",r"\bgraphic design\b",r"\bdesign creatives?\b"],
"TikTok Ads":[r"\btiktok ads?\b"],
"SEO":[r"\bseo\b",r"\bsearch engine optimization\b"],
"A/B Testing":[r"\ba/?b testing\b",r"\bsplit testing\b"],
}

MUST_HEADINGS=("must have","requirements","required","essential","minimum qualifications","qualifications","what you need","skills required")
PREF_HEADINGS=("nice to have","preferred","desirable","bonus","plus","advantage")


def clean_line(s): return re.sub(r"\s+"," ",(s or "")).strip()

def heading_kind(line):
    x=clean_line(line).lower().rstrip(":-")
    if any(h in x for h in PREF_HEADINGS): return "preferred"
    if any(h in x for h in MUST_HEADINGS): return "required"
    return None

def sectionize_jd(jd):
    """Return blocks tagged required/preferred/general, preserving the heading as provenance."""
    lines=[clean_line(x) for x in (jd or "").splitlines() if clean_line(x)]
    blocks=[]; current_kind="general"; current_heading="General JD text"; bucket=[]
    for line in lines:
        k=heading_kind(line)
        if k:
            if bucket: blocks.append((current_kind,current_heading,"\n".join(bucket)))
            # Support inline headings such as "Must Have: Meta Ads, Google Ads".
            parts=re.split(r"[:\-]",line,maxsplit=1)
            current_kind=k; current_heading=parts[0].strip() if parts else line.rstrip(":")
            bucket=[]
            if len(parts)==2 and clean_line(parts[1]): bucket.append(clean_line(parts[1]))
        else:
            bucket.append(line)
    if bucket: blocks.append((current_kind,current_heading,"\n".join(bucket)))
    if not blocks and jd.strip(): blocks=[("general","General JD text",jd)]
    return blocks

def find_in(text, patterns):
    return any(re.search(p,text,re.I) for p in patterns)

def requirement_rows(jd):
    blocks=sectionize_jd(jd)
    rows=[]; seen={}
    # Prefer explicit sections. General JD text is used only when there are no explicit sections for a skill.
    for kind, heading, body in blocks:
        for name,pats in SKILLS.items():
            if find_in(body,pats):
                candidate={"name":name,"patterns":pats,"category":kind,"source":heading,"source_text":body,"type":"skill"}
                priority={"required":3,"preferred":2,"general":1}
                old=seen.get(name)
                if old is None or priority[kind]>priority[old["category"]]: seen[name]=candidate
    rows=list(seen.values())
    # Years: search within the block that actually contains it; required beats preferred.
    year_seen=None
    for kind, heading, body in blocks:
        m=re.search(r"\b(\d+)\s*\+?\s*years?(?:\s+of)?(?:\s+relevant)?",body,re.I)
        if m:
            cand={"name":f"{m.group(1)}+ years relevant experience","years":int(m.group(1)),"category":kind,"source":heading,"source_text":body,"type":"experience"}
            if year_seen is None or {"general":1,"preferred":2,"required":3}[kind]>{"general":1,"preferred":2,"required":3}[year_seen["category"]]: year_seen=cand
    if year_seen: rows.append(year_seen)
    # Location is a suitability preference unless the JD explicitly says on-site / hybrid in a city.
    for kind,heading,body in blocks:
        low=body.lower()
        for city in ["lagos","abuja","port harcourt","ibadan","enugu"]:
            if city in low:
                cat="required" if kind=="required" and ("hybrid" in low or "on-site" in low or "onsite" in low) else "preferred"
                rows.append({"name":f"{city.title()} location / role suitability","location":city,"category":cat,"source":heading,"source_text":body,"type":"location"})
                break
        if any(city in low for city in ["lagos","abuja","port harcourt","ibadan","enugu"]): break
    # Domain experience only if actually stated in JD.
    full=jd.lower()
    if any(x in full for x in ["fintech","banking","financial services"]):
        source="General JD text"; cat="preferred"
        for kind,heading,body in blocks:
            if any(x in body.lower() for x in ["fintech","banking","financial services"]):
                source=heading; cat="required" if kind=="required" else "preferred"; break
        rows.append({"name":"Fintech / banking experience","patterns":[r"\bfintech\b",r"\bbank(?:ing)?\b",r"\bfinancial services\b",r"\bkuda\b",r"\bopay\b",r"\bcarbon\b",r"\bflutterwave\b",r"\bgtbank\b",r"\bmoniepoint\b"],"category":cat,"source":source,"source_text":"fintech / banking","type":"preference"})
    # Deduplicate by name, keeping strongest category.
    out={}
    pri={"required":3,"preferred":2,"general":1}
    for r in rows:
        if r["name"] not in out or pri[r["category"]]>pri[out[r["name"]]["category"]]: out[r["name"]]=r
    return list(out.values())

# ------------------------- EVIDENCE VALIDATION --------------------------
NEG_PATTERNS=[
    r"\bno\s+(?:experience(?:\s+with|\s+in)?|knowledge\s+of|background\s+in|[a-z0-9/&+ .-]+)",
    r"\bwithout\s+(?:experience(?:\s+with|\s+in)?|[a-z0-9/&+ .-]+)",
    r"\black(?:s|ing)?\s+(?:experience(?:\s+with|\s+in)?|[a-z0-9/&+ .-]+)",
    r"\bnever\s+(?:used|worked with|managed)\b",
    r"\bnot\s+(?:experienced|proficient|familiar)\b",
]
WEAK_TERMS=("basic","assisted","support","supported","boosted posts","sometimes","familiar","exposure","intern","trainee","only")
STRONG_TERMS=("managed","led","owned","budget","certified","implemented","optimized","built","reduced","increased","conversion","campaign","delivered","drove","grew","improved")

def split_evidence_units(text):
    text=re.sub(r"\r", "\n", text or "")
    # Keep bullet/line statements intact; also split sentences.
    units=[]
    for line in text.split("\n"):
        line=clean_line(line)
        if not line: continue
        units.extend([clean_line(x) for x in re.split(r"(?<=[.!?;])\s+",line) if clean_line(x)])
    return units

def is_explicit_negative(unit, match_span):
    low=unit.lower()
    # A GAPS / skills-gap label plus negation is particularly strong.
    left=low[max(0,match_span[0]-80):match_span[1]+80]
    if re.search(r"\b(gaps?|missing|lacks?)\s*[:\-]",low) and re.search(r"\b(no|without|lack(?:s|ing)?)\b",left): return True
    # Check negation in a local window immediately before the matched skill.
    prefix=low[max(0,match_span[0]-50):match_span[0]]
    if re.search(r"\b(no|without|lacks?|lack of|never used|never worked with|not experienced|not familiar with)\b",prefix): return True
    # Explicit 'No X' where X is the unit's matched skill; conservative window after no.
    if re.search(r"\bno\b.{0,35}",prefix+low[match_span[0]:match_span[1]],re.I): return True
    return False

def evidence_for(text, patterns):
    positives=[]; negatives=[]
    for unit in split_evidence_units(text):
        for pat in patterns:
            m=re.search(pat,unit,re.I)
            if not m: continue
            if is_explicit_negative(unit,m.span()): negatives.append(unit); break
            low=unit.lower()
            if any(term in low for term in WEAK_TERMS): positives.append((0.35,"Weak / limited evidence",unit))
            elif any(term in low for term in STRONG_TERMS): positives.append((1.0,"Direct evidence",unit))
            else: positives.append((0.70,"Evidence present",unit))
            break
    if positives:
        positives.sort(key=lambda x:x[0],reverse=True)
        return positives[0][0],positives[0][1],positives[0][2],"Positive evidence"
    if negatives:
        return 0.0,"Explicitly absent",negatives[0],"Negative evidence"
    return 0.0,"No evidence found","","No evidence"

def years(text):
    current=date.today().year; total=0
    ranges=re.findall(r"\b(19\d{2}|20\d{2})\s*(?:-|–|to)\s*(19\d{2}|20\d{2}|present|current)\b",(text or "").lower())
    for a,b in ranges:
        end=current if b in ["present","current"] else int(b)
        total += max(0,end-int(a))
    return total

def candidate_name(filename,text):
    for line in [x.strip() for x in text.splitlines() if x.strip()][:12]:
        words=line.split()
        if 2<=len(words)<=5 and len(line)<60 and line.upper()==line and not any(w in line.upper() for w in ["EXPERIENCE","SKILLS","EDUCATION","CURRICULUM","VITAE"]):
            return line.title()
    return Path(filename).stem.replace("_"," ").replace("-"," ").strip()

# ------------------------------ SCORING ---------------------------------
def classify_category(cat):
    return "Must-have" if cat=="required" else "Preferred" if cat=="preferred" else "General JD mention"

def score_candidate(name,text,reqs):
    audit=[]; mandatory=[]; preferred=[]
    for r in reqs:
        status=""; evidence=""; evtype=""
        if r["type"] in ["skill","preference"]:
            level,status,evidence,evtype=evidence_for(text,r["patterns"])
        elif r["type"]=="experience":
            y=years(text); target=r["years"]; level=min(1.0,y/target) if target else 0
            status="Meets requirement" if y>=target else "Below requirement"
            evidence=f"Estimated dated experience: {y} year(s)"
            evtype="Positive evidence" if y else "No evidence"
        else:
            low=text.lower(); city=r["location"]
            if city in low: level=1.0; status="Location matches"; evidence=city.title(); evtype="Positive evidence"
            elif "remote" in low: level=0.5; status="Remote / review"; evidence="Remote location stated"; evtype="Partial evidence"
            else: level=0.0; status="No matching location evidence"; evidence=""; evtype="No evidence"
        row={"Requirement":r["name"],"Category":classify_category(r["category"]),"Source / Provenance":r["source"],"Evidence Type":evtype,"Evidence Level":round(level,2),"Status":status,"Evidence":evidence}
        if r["category"]=="required": mandatory.append(row)
        elif r["category"]=="preferred": preferred.append(row)
        else: preferred.append(row) # General mentions never count as mandatory.
        audit.append(row)
    mandatory_score=(sum(x["Evidence Level"] for x in mandatory)/len(mandatory)*100) if mandatory else 100.0
    preferred_bonus=(sum(x["Evidence Level"] for x in preferred)/len(preferred)*10) if preferred else 0.0
    final=mandatory_score + preferred_bonus
    # Missing mandatory requirements trigger a transparent cap.
    missing=[x for x in mandatory if x["Evidence Level"]==0]
    if missing: final=min(final, max(0, 94-6*(len(missing)-1)))
    final=round(min(100,final))
    for row in audit:
        row["Points Earned"] = round(row["Evidence Level"]*(100/len(mandatory)),2) if row in mandatory and mandatory else round(row["Evidence Level"]*(10/len(preferred)),2) if row in preferred and preferred else 0
    group="Excellent" if final>=90 else "Good" if final>=70 else "Moderate" if final>=50 else "Do Not Hire"
    return final,group,audit,mandatory_score,preferred_bonus

# ------------------------------- REPORTS --------------------------------
def charts(df):
    d=Path(tempfile.mkdtemp()); files=[]
    c=df["Ranking Group"].value_counts(); fig,ax=plt.subplots(figsize=(7,4.5)); ax.pie(c.values,labels=c.index,autopct="%1.0f%%"); ax.set_title("Candidate Ranking Distribution"); p=d/"rank.png"; fig.savefig(p,dpi=150,bbox_inches="tight"); plt.close(fig); files.append(p)
    fig,ax=plt.subplots(figsize=(7,4.5)); ax.hist(df["Fit %"],bins=min(10,max(3,len(df))),edgecolor="black"); ax.set_xlabel("Fit Score (%)"); ax.set_ylabel("Number of Candidates"); ax.set_title("Score Distribution"); p=d/"scores.png"; fig.savefig(p,dpi=150,bbox_inches="tight"); plt.close(fig); files.append(p)
    top=df.head(10).iloc[::-1]; fig,ax=plt.subplots(figsize=(7,4.8)); ax.barh(top["Name"],top["Fit %"]); ax.set_xlim(0,100); ax.set_xlabel("Fit Score (%)"); ax.set_title("Top 10 Candidates"); p=d/"top.png"; fig.savefig(p,dpi=150,bbox_inches="tight"); plt.close(fig); files.append(p); return files

def pdf_report(df,client,address,officer,email,project,audits,reqs):
    bio=io.BytesIO(); doc=SimpleDocTemplate(bio,pagesize=A4,rightMargin=28,leftMargin=28,topMargin=30,bottomMargin=30); stl=getSampleStyleSheet(); story=[]
    story += [Paragraph("CONFIDENTIAL CANDIDATE SCREENING REPORT",stl["Title"]),Spacer(1,10)]
    for label,val in [("Prepared exclusively for",client),("Client Address",address),("Client Contact Officer",officer),("Client Contact Email",email),("Recruitment Project",project),("Screening date",date.today().isoformat())]:
        if val: story.append(Paragraph(f"<b>{label}:</b> {val}",stl["BodyText"]))
    story += [Spacer(1,10),Paragraph("Executive Summary",stl["Heading1"]),Paragraph(f"{len(df)} candidates were screened, scored and ranked against the supplied Job Description.",stl["BodyText"])]
    data=[["Ranking Group","Candidates"]]+[[k,int(v)] for k,v in df["Ranking Group"].value_counts().items()]; t=Table(data); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey)])); story += [Spacer(1,8),t,Spacer(1,10),Paragraph("Top Recommended Candidates",stl["Heading1"])]
    for _,r in df.head(10).iterrows(): story.append(Paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})",stl["BodyText"]))
    story += [Spacer(1,10),Paragraph("Scoring Method",stl["Heading2"]),Paragraph("Must-have requirements form the primary score. Preferred requirements can add a limited bonus. Explicit negative statements are not treated as positive evidence. Each requirement shows its source in the Job Description.",stl["BodyText"])]
    for p in charts(df): story += [PageBreak(),Image(str(p),width=480,height=310)]
    story += [PageBreak(),Paragraph("Candidate-by-Candidate Scoring Audit",stl["Heading1"])]
    for name in df["Name"]:
        story += [Paragraph(name,stl["Heading2"])]
        data=[["Requirement","Category","Source","Evidence Type","Status","Evidence"]]
        for x in audits[name]: data.append([x["Requirement"],x["Category"],x["Source / Provenance"][:50],x["Evidence Type"],x["Status"],x["Evidence"][:110]])
        t=Table(data,colWidths=[85,55,65,65,75,150],repeatRows=1); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.3,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),6.5)])); story += [t,Spacer(1,8)]
    story += [Spacer(1,10),Paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener.",stl["Italic"])]
    doc.build(story); return bio.getvalue()

def docx_report(df,client,address,officer,email,project,audits):
    doc=Document(); doc.add_heading("CONFIDENTIAL CANDIDATE SCREENING REPORT",0)
    for label,val in [("Prepared exclusively for",client),("Client Address",address),("Client Contact Officer",officer),("Client Contact Email",email),("Recruitment Project",project),("Screening date",date.today().isoformat())]:
        if val: doc.add_paragraph(f"{label}: {val}")
    doc.add_heading("Executive Summary",1); doc.add_paragraph(f"{len(df)} candidates were screened, scored and ranked against the supplied Job Description.")
    doc.add_heading("Top Recommended Candidates",1)
    for _,r in df.head(10).iterrows(): doc.add_paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})")
    doc.add_heading("Scoring Method",1); doc.add_paragraph("Must-have requirements form the primary score. Preferred requirements add only a limited bonus. Explicit negative statements are not treated as positive evidence. Every audit row records the Job Description source.")
    doc.add_page_break(); doc.add_heading("Candidate-by-Candidate Scoring Audit",0)
    for name in df["Name"]:
        doc.add_heading(name,1); table=doc.add_table(rows=1,cols=6); table.style="Table Grid"
        for i,h in enumerate(["Requirement","Category","Source","Evidence Type","Status","Evidence"]): table.rows[0].cells[i].text=h
        for x in audits[name]:
            vals=[x["Requirement"],x["Category"],x["Source / Provenance"],x["Evidence Type"],x["Status"],x["Evidence"]]; cells=table.add_row().cells
            for i,v in enumerate(vals): cells[i].text=str(v)
    doc.add_paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener."); bio=io.BytesIO(); doc.save(bio); return bio.getvalue()

def workbook(df,audits,dups,reqs):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name="Master Ranking")
        for g in ["Excellent","Good","Moderate","Do Not Hire"]: df[df["Ranking Group"]==g].to_excel(w,index=False,sheet_name=g[:31])
        rows=[]
        for name,a in audits.items():
            for x in a: rows.append({"Candidate":name,**x})
        pd.DataFrame(rows).to_excel(w,index=False,sheet_name="Scoring Audit")
        pd.DataFrame([{k:v for k,v in r.items() if k not in ["patterns","source_text"]} for r in reqs]).to_excel(w,index=False,sheet_name="JD Requirements")
        pd.DataFrame(dups).to_excel(w,index=False,sheet_name="Duplicates")
    return bio.getvalue()

# -------------------------------- UI -----------------------------------
st.title("🎯 Opportunity Hub CV Screener v2.7")
st.caption("Evidence Validation & Requirement Classification — prevents false-positive skill matches and records JD provenance.")
with st.expander("Client and recruitment details",expanded=True):
    a,b=st.columns(2)
    client=a.text_input("Client name")
    project=b.text_input("Recruitment role / job")
    address=a.text_input("Client address")
    officer=b.text_input("Client contact officer")
    email=a.text_input("Client contact email")
st.subheader("Job Description")
jd_text=st.text_area("Paste Job Description (optional)",height=180)
jd_files=st.file_uploader("Upload Job Description file",type=["pdf","docx","txt","csv","xlsx","xls"],accept_multiple_files=True)
st.subheader("CV Intake")
cv_files=st.file_uploader("Upload individual CV files or ZIP batches",type=["pdf","docx","txt","csv","xlsx","xls","zip"],accept_multiple_files=True)
st.caption("PDF support uses PyMuPDF first and pypdf fallback. Scanned/image-only PDFs are flagged instead of silently treated as readable.")
if "results_v27" not in st.session_state: st.session_state.results_v27=None
if st.button("Screen CVs",type="primary"):
    errors=[]; final=jd_text.strip()
    for f in jd_files or []:
        tx,_,er=extract(f.name,f.getvalue()); final += "\n"+tx if tx else ""; errors += [f"JD {f.name}: {er}"] if er else []
    if len(final.strip())<30: st.error("A readable Job Description is required. The PDF may be scanned/image-only."); st.stop()
    reqs=requirement_rows(final)
    if not reqs: st.error("No usable requirements could be extracted from the Job Description. Use clear Must Have / Required / Nice to Have wording."); st.stop()
    records=[]
    for f in cv_files or []:
        for name,data in unpack(f.name,f.getvalue()):
            tx,method,er=extract(name,data); records.append({"filename":name,"text":tx,"hash":sha256(data),"error":er})
    if not records: st.error("Upload at least one CV file or ZIP batch."); st.stop()
    seen={}; sigs={}; dups=[]; valid=[]
    for r in records:
        if r["hash"] in seen: dups.append({"File":r["filename"],"Duplicate Type":"Exact file duplicate","Duplicate Of":seen[r["hash"]]}); continue
        sig=re.sub(r"\W+","",r["text"].lower())[:5000]
        if sig and sig in sigs: dups.append({"File":r["filename"],"Duplicate Type":"Duplicate extracted content","Duplicate Of":sigs[sig]}); continue
        seen[r["hash"]]=r["filename"]; sigs[sig]=r["filename"] if sig else r["filename"]
        if r["error"]: errors.append(f"CV {r['filename']}: {r['error']}")
        else: valid.append(r)
    rows=[]; audits={}
    for r in valid:
        name=candidate_name(r["filename"],r["text"]); s,g,audit,mandatory_score,preferred_bonus=score_candidate(name,r["text"],reqs)
        missing=[x["Requirement"] for x in audit if x["Category"]=="Must-have" and x["Evidence Level"]==0]
        rows.append({"Name":name,"Fit %":s,"Ranking Group":g,"Mandatory Score":round(mandatory_score,1),"Preferred Bonus":round(preferred_bonus,1),"Years Exp":years(r["text"]),"File":r["filename"],"Why Not 100%":"; ".join(missing[:4]) if missing else "No material must-have gaps detected"}); audits[name]=audit
    df=pd.DataFrame(rows).sort_values(["Fit %","Years Exp"],ascending=[False,False]).reset_index(drop=True) if rows else pd.DataFrame(columns=["Name","Fit %","Ranking Group","Mandatory Score","Preferred Bonus","Years Exp","File","Why Not 100%"])
    st.session_state.results_v27={"df":df,"audits":audits,"duplicates":dups,"errors":errors,"client":client,"address":address,"officer":officer,"email":email,"project":project,"reqs":reqs}
res=st.session_state.results_v27
if res:
    df=res["df"]; audits=res["audits"]; st.success(f"Screening complete: {len(df)} unique readable CVs screened.")
    if res["errors"]: st.warning("Some files need attention:\n\n"+"\n\n".join(res["errors"]))
    st.subheader("Job Description Requirements Detected")
    req_display=pd.DataFrame([{k:v for k,v in r.items() if k not in ["patterns","source_text"]} for r in res["reqs"]])
    st.dataframe(req_display,use_container_width=True)
    if res["duplicates"]: st.info(f"{len(res['duplicates'])} duplicate entries excluded from scoring."); st.dataframe(pd.DataFrame(res["duplicates"]),use_container_width=True)
    st.subheader("Master Ranking"); st.dataframe(df,use_container_width=True)
    st.subheader("Candidate-by-Candidate Scoring Audit")
    for name in df["Name"]:
        with st.expander(f"{name} — evidence, negatives and JD source"): st.dataframe(pd.DataFrame(audits[name]),use_container_width=True)
    if not df.empty:
        ch=charts(df); c1,c2=st.columns(2); c1.image(str(ch[0]),use_container_width=True); c2.image(str(ch[1]),use_container_width=True); st.image(str(ch[2]),use_container_width=True)
        st.subheader("Downloads")
        st.download_button("Download professional workbook",workbook(df,audits,res["duplicates"],res["reqs"]),"candidate_screening_workbook.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Download CSV",df.to_csv(index=False).encode(),"candidate_screening.csv","text/csv")
        st.download_button("Download professional PDF report",pdf_report(df,res["client"],res["address"],res["officer"],res["email"],res["project"],audits,res["reqs"]),"candidate_screening_report.pdf","application/pdf")
        st.download_button("Download professional DOCX report",docx_report(df,res["client"],res["address"],res["officer"],res["email"],res["project"],audits),"candidate_screening_report.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document")
