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

st.set_page_config(page_title="Opportunity Hub CV Screener v2.6", page_icon="🎯", layout="wide")

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

def unpack(name,data):
    if Path(name).suffix.lower()!=".zip": return [(name,data)]
    out=[]
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for i in z.infolist():
                if not i.is_dir() and Path(i.filename).suffix.lower() in [".pdf",".docx",".txt",".csv",".xlsx",".xls"]:
                    out.append((Path(i.filename).name,z.read(i)))
    except Exception as e: return [(name,b"")]
    return out

def norm(s): return re.sub(r"\s+"," ",(s or "").lower()).strip()

def sentence(text,patterns):
    for s in re.split(r"(?<=[\.\!\?;\n])\s+", text or ""):
        if any(re.search(p,s,re.I) for p in patterns): return s.strip()[:260]
    return ""

SKILLS={
"Meta Ads":[r"\bmeta ads?\b",r"\bfacebook ads?\b",r"\binstagram ads?\b"],
"Google Ads":[r"\bgoogle ads?\b",r"\badwords\b"],
"Email Marketing":[r"\bemail campaigns?\b",r"\bemail marketing\b",r"\bnewsletter[s]?\b"],
"Copywriting":[r"\bcopywriting\b",r"\bwrite copy\b",r"\bad copy\b",r"\bsocial copy\b"],
"GA4 / Google Analytics":[r"\bga4\b",r"\bgoogle analytics\b"],
"Canva":[r"\bcanva\b"],"HubSpot":[r"\bhubspot\b"],
"Basic Design":[r"\bbasic design\b",r"\bgraphic design\b",r"\bdesign creatives?\b"],
"TikTok Ads":[r"\btiktok ads?\b"],"SEO":[r"\bseo\b",r"\bsearch engine optimization\b"],
"A/B Testing":[r"\ba/?b testing\b",r"\bsplit testing\b"]}

def jd_requirements(jd):
    j=norm(jd); req=[]
    for name,pats in SKILLS.items():
        ev=sentence(jd,pats)
        if ev:
            nice=any(x in ev.lower() for x in ["nice to have","preferred","bonus","advantage","plus"])
            req.append({"name":name,"patterns":pats,"required":not nice,"weight":4 if not nice else 1,"type":"skill"})
    m=re.search(r"(\d+)\s*\+?\s*years?",j)
    if m: req.append({"name":f"{m.group(1)}+ years relevant experience","years":int(m.group(1)),"required":True,"weight":4,"type":"experience"})
    for city in ["lagos","abuja","port harcourt","ibadan","enugu"]:
        if city in j:
            req.append({"name":f"{city.title()} location / role suitability","location":city,"required":False,"weight":2,"type":"location"}); break
    if any(x in j for x in ["fintech","banking","financial services"]):
        req.append({"name":"Fintech / banking experience","patterns":[r"\bfintech\b",r"\bbank\b",r"\bbanking\b",r"\bkuda\b",r"\bopay\b",r"\bcarbon\b",r"\bflutterwave\b",r"\bgtbank\b",r"\bmoniepoint\b"],"required":False,"weight":2,"type":"preference"})
    return req

def years(text):
    ranges=re.findall(r"\b(20\d{2})\s*(?:-|–|to)\s*(20\d{2}|present|current)\b",(text or "").lower())
    if ranges:
        return sum(max(0,(date.today().year if b in ["present","current"] else int(b))-int(a)) for a,b in ranges)
    ys=[int(x) for x in re.findall(r"\b(19\d{2}|20\d{2})\b",text or "")]
    return max(0,max(ys)-min(ys)) if ys else 0

def level(text,pats):
    ev=sentence(text,pats)
    if not ev: return 0,"No evidence found",""
    e=ev.lower()
    if any(x in e for x in ["basic","assisted","support","boosted posts","sometimes","familiar"]): return .35,"Weak / limited evidence",ev
    if any(x in e for x in ["managed","led","owned","budget","certified","implemented","optimized","built","reduced","increased","conversion","campaign"]): return 1,"Direct evidence",ev
    return .7,"Evidence present",ev

def candidate_name(filename,text):
    for line in [x.strip() for x in text.splitlines() if x.strip()][:12]:
        if 2<=len(line.split())<=5 and len(line)<60 and line.upper()==line: return line.title()
    return Path(filename).stem.replace("_"," ").replace("-"," ").strip()

def score(name,text,reqs):
    audit=[]; earned=possible=0
    for r in reqs:
        w=r["weight"]; possible+=w; pts=0; status=""; evidence=""
        if r["type"] in ["skill","preference"]:
            lv,status,evidence=level(text,r["patterns"]); pts=w*lv
        elif r["type"]=="experience":
            y=years(text); pts=w*min(1,y/r["years"]); status="Meets requirement" if y>=r["years"] else "Below requirement"; evidence=f"Estimated relevant experience: {y} year(s)"
        else:
            t=norm(text); city=r["location"]
            if city in t: pts=w; status="Location matches"; evidence=city.title()
            elif "remote" in t: pts=w*.5; status="Remote / review"; evidence="Remote location stated"
            else: pts=w*.25; status="Location review required"; evidence="No clear matching location evidence"
        earned+=pts; audit.append({"Requirement":r["name"],"Required":"Yes" if r.get("required") else "No / Preferred","Weight":w,"Points Earned":round(pts,2),"Status":status,"Evidence":evidence})
    s=round(earned/possible*100) if possible else 0
    missing=sum(1 for x in audit if x["Required"]=="Yes" and x["Points Earned"]==0)
    if missing: s=min(s,94-min(20,missing*4))
    group="Excellent" if s>=90 else "Good" if s>=70 else "Moderate" if s>=50 else "Do Not Hire"
    return s,group,audit

def charts(df):
    d=Path(tempfile.mkdtemp()); files=[]
    c=df["Ranking Group"].value_counts(); fig,ax=plt.subplots(figsize=(7,4.5)); ax.pie(c.values,labels=c.index,autopct="%1.0f%%"); ax.set_title("Candidate Ranking Distribution"); p=d/"rank.png"; fig.savefig(p,dpi=150,bbox_inches="tight"); plt.close(fig); files.append(p)
    fig,ax=plt.subplots(figsize=(7,4.5)); ax.hist(df["Fit %"],bins=min(10,max(3,len(df))),edgecolor="black"); ax.set_xlabel("Fit Score (%)"); ax.set_ylabel("Number of Candidates"); ax.set_title("Score Distribution"); p=d/"scores.png"; fig.savefig(p,dpi=150,bbox_inches="tight"); plt.close(fig); files.append(p)
    top=df.head(10).iloc[::-1]; fig,ax=plt.subplots(figsize=(7,4.8)); ax.barh(top["Name"],top["Fit %"]); ax.set_xlim(0,100); ax.set_xlabel("Fit Score (%)"); ax.set_title("Top 10 Candidates"); p=d/"top.png"; fig.savefig(p,dpi=150,bbox_inches="tight"); plt.close(fig); files.append(p); return files

def pdf_report(df,client,project,audits):
    bio=io.BytesIO(); doc=SimpleDocTemplate(bio,pagesize=A4,rightMargin=36,leftMargin=36,topMargin=36,bottomMargin=36); stl=getSampleStyleSheet(); story=[]
    story += [Paragraph("CONFIDENTIAL CANDIDATE SCREENING REPORT",stl["Title"]),Spacer(1,12),Paragraph(f"<b>Prepared exclusively for:</b> {client or 'Client'}",stl["BodyText"]),Paragraph(f"<b>Recruitment project:</b> {project or 'Recruitment Screening'}",stl["BodyText"]),Paragraph(f"<b>Screening date:</b> {date.today().isoformat()}",stl["BodyText"]),Spacer(1,12),Paragraph("Executive Summary",stl["Heading1"]),Paragraph(f"{len(df)} candidates were screened, scored and ranked against the supplied Job Description.",stl["BodyText"])]
    data=[["Ranking Group","Candidates"]]+[[k,int(v)] for k,v in df["Ranking Group"].value_counts().items()]; t=Table(data); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey)])); story += [Spacer(1,8),t,Spacer(1,12),Paragraph("Top Recommended Candidates",stl["Heading1"])]
    for _,r in df.head(10).iterrows(): story.append(Paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})",stl["BodyText"]))
    for p in charts(df): story += [PageBreak(),Image(str(p),width=480,height=310)]
    story += [PageBreak(),Paragraph("Candidate-by-Candidate Audit Trail",stl["Heading1"])]
    for name in df["Name"]:
        story += [Paragraph(name,stl["Heading2"])]
        data=[["Requirement","Required","Points","Status","Evidence"]]
        for x in audits[name]: data.append([x["Requirement"],x["Required"],str(x["Points Earned"]),x["Status"],x["Evidence"][:110]])
        t=Table(data,colWidths=[105,50,40,90,220],repeatRows=1); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.3,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),7)])); story += [t,Spacer(1,8)]
    story += [Paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener.",stl["Italic"])]
    doc.build(story); return bio.getvalue()

def docx_report(df,client,project,audits):
    doc=Document(); doc.add_heading("CONFIDENTIAL CANDIDATE SCREENING REPORT",0); doc.add_paragraph(f"Prepared exclusively for: {client or 'Client'}"); doc.add_paragraph(f"Recruitment project: {project or 'Recruitment Screening'}"); doc.add_paragraph(f"Screening date: {date.today().isoformat()}"); doc.add_heading("Executive Summary",1); doc.add_paragraph(f"{len(df)} candidates were screened, scored and ranked against the supplied Job Description."); doc.add_heading("Top Recommended Candidates",1)
    for _,r in df.head(10).iterrows(): doc.add_paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})")
    doc.add_page_break(); doc.add_heading("Candidate-by-Candidate Audit Trail",0)
    for name in df["Name"]:
        doc.add_heading(name,1); table=doc.add_table(rows=1,cols=5); table.style="Table Grid"
        for i,h in enumerate(["Requirement","Required","Points","Status","Evidence"]): table.rows[0].cells[i].text=h
        for x in audits[name]:
            vals=[x["Requirement"],x["Required"],str(x["Points Earned"]),x["Status"],x["Evidence"]]; cells=table.add_row().cells
            for i,v in enumerate(vals): cells[i].text=str(v)
    doc.add_paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener."); bio=io.BytesIO(); doc.save(bio); return bio.getvalue()

def workbook(df,audits,dups):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name="Master Ranking")
        for g in ["Excellent","Good","Moderate","Do Not Hire"]: df[df["Ranking Group"]==g].to_excel(w,index=False,sheet_name=g[:31])
        rows=[]
        for name,a in audits.items():
            for x in a: rows.append({"Candidate":name,**x})
        pd.DataFrame(rows).to_excel(w,index=False,sheet_name="Scoring Audit"); pd.DataFrame(dups).to_excel(w,index=False,sheet_name="Duplicates")
    return bio.getvalue()

st.title("🎯 Opportunity Hub CV Screener v2.6")
st.caption("Reliability & Auditability Fix — robust PDF intake, evidence-based scoring and full scoring audit trail.")
with st.expander("Client and recruitment details",expanded=True):
    a,b=st.columns(2); client=a.text_input("Client name"); project=b.text_input("Recruitment role / job"); a.text_input("Client address"); b.text_input("Client contact officer"); a.text_input("Client contact email")
st.subheader("Job Description")
jd_text=st.text_area("Paste Job Description (optional)",height=180)
jd_files=st.file_uploader("Upload Job Description file",type=["pdf","docx","txt","csv","xlsx","xls"],accept_multiple_files=True)
st.subheader("CV Intake")
cv_files=st.file_uploader("Upload individual CV files or ZIP batches",type=["pdf","docx","txt","csv","xlsx","xls","zip"],accept_multiple_files=True)
st.caption("PDF support uses PyMuPDF first and pypdf fallback. Scanned/image-only PDFs are clearly flagged instead of silently failing.")
if "results" not in st.session_state: st.session_state.results=None
if st.button("Screen CVs",type="primary"):
    errors=[]; final=jd_text.strip()
    for f in jd_files or []:
        tx,_,er=extract(f.name,f.getvalue()); final += "\n"+tx if tx else ""; errors += [f"JD {f.name}: {er}"] if er else []
    if len(final.strip())<30: st.error("A readable Job Description is required. The PDF may be scanned/image-only."); st.stop()
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
    reqs=jd_requirements(final); rows=[]; audits={}
    for r in valid:
        name=candidate_name(r["filename"],r["text"]); s,g,audit=score(name,r["text"],reqs); missing=[x["Requirement"] for x in audit if x["Required"]=="Yes" and x["Points Earned"]==0]
        rows.append({"Name":name,"Fit %":s,"Ranking Group":g,"Years Exp":years(r["text"]),"File":r["filename"],"Why Not 100%":"; ".join(missing[:4]) if missing else "No material required gaps detected"}); audits[name]=audit
    df=pd.DataFrame(rows).sort_values(["Fit %","Years Exp"],ascending=[False,False]).reset_index(drop=True) if rows else pd.DataFrame(columns=["Name","Fit %","Ranking Group","Years Exp","File","Why Not 100%"])
    st.session_state.results={"df":df,"audits":audits,"duplicates":dups,"errors":errors,"client":client,"project":project}
res=st.session_state.results
if res:
    df=res["df"]; audits=res["audits"]; st.success(f"Screening complete: {len(df)} unique readable CVs screened.")
    if res["errors"]: st.warning("Some files need attention:\n\n"+"\n\n".join(res["errors"]))
    if res["duplicates"]: st.info(f"{len(res['duplicates'])} duplicate entries excluded from scoring."); st.dataframe(pd.DataFrame(res["duplicates"]),use_container_width=True)
    st.subheader("Master Ranking"); st.dataframe(df,use_container_width=True)
    st.subheader("Scoring Audit Trail")
    for name in df["Name"]:
        with st.expander(f"{name} — evidence and points"): st.dataframe(pd.DataFrame(audits[name]),use_container_width=True)
    if not df.empty:
        ch=charts(df); c1,c2=st.columns(2); c1.image(str(ch[0]),use_container_width=True); c2.image(str(ch[1]),use_container_width=True); st.image(str(ch[2]),use_container_width=True)
        st.subheader("Downloads"); st.download_button("Download professional workbook",workbook(df,audits,res["duplicates"]),"candidate_screening_workbook.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); st.download_button("Download CSV",df.to_csv(index=False).encode(),"candidate_screening.csv","text/csv"); st.download_button("Download professional PDF report",pdf_report(df,res["client"],res["project"],audits),"candidate_screening_report.pdf","application/pdf"); st.download_button("Download professional DOCX report",docx_report(df,res["client"],res["project"],audits),"candidate_screening_report.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document")
