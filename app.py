import io, re, zipfile, datetime
from pathlib import Path
from urllib.parse import urlparse
import pandas as pd
import requests
import streamlit as st
import matplotlib.pyplot as plt
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

st.set_page_config(page_title="Opportunity Hub Screener v2.4", layout="wide")
SUPPORTED = [".pdf", ".docx", ".txt"]

def detect_extension(filename, data=b"", content_type=""):
    ext = Path(filename).suffix.lower()
    if ext in SUPPORTED: return ext
    if data.startswith(b"%PDF"): return ".pdf"
    if data.startswith(b"PK\x03\x04"): return ".docx"
    ct = (content_type or "").lower()
    if "pdf" in ct: return ".pdf"
    if "word" in ct or "officedocument.wordprocessingml" in ct: return ".docx"
    return ".txt"

def read_document(filename, data, content_type=""):
    ext = detect_extension(filename, data, content_type)
    if ext == ".pdf":
        return "\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(data)).pages)
    if ext == ".docx":
        return "\n".join(p.text for p in Document(io.BytesIO(data)).paragraphs)
    return data.decode("utf-8", errors="ignore")

def drive_url(url):
    m = re.search(r"(?:/d/|id=)([\w-]{10,})", url)
    if "drive.google.com" in url and m:
        return f"https://drive.usercontent.google.com/download?id={m.group(1)}&export=download&confirm=t"
    return url

def download_public_file(url):
    r = requests.get(drive_url(url), headers={"User-Agent":"Mozilla/5.0"}, timeout=120, allow_redirects=True)
    r.raise_for_status()
    data = r.content
    if data[:100].lower().startswith(b"<!doctype html") or b"<html" in data[:500].lower():
        raise ValueError("Link returned a web page, not the file. Set sharing to Anyone with the link.")
    cd = r.headers.get("content-disposition","")
    m = re.search(r'filename[^=]*=(?:UTF-8\'\')?["\']?([^"\';]+)', cd, re.I)
    name = m.group(1).strip() if m else Path(urlparse(url).path).name or "document"
    return name, data, r.headers.get("content-type","")

def extract_name(text, fallback):
    lines=[re.sub(r"\s+"," ",x).strip() for x in text.splitlines() if x.strip()]
    blocked=["curriculum","resume","digital marketing lead","performance marketing","email & content","social media manager","marketing intern","sales executive","graphic designer","fresh graduate","hr generalist","mechanical engineer"]
    for line in lines[:10]:
        if 2 <= len(line.split()) <= 5 and len(line)<70 and not any(x in line.lower() for x in blocked) and not re.search(r"[\d@]",line):
            return line.title()
    s=re.sub(r"^cv[_\s-]*\d+[_\s-]*","",Path(fallback).stem,flags=re.I)
    return s.replace("_"," ").replace("-"," ").title()

def years_experience(text):
    ranges=re.findall(r"(20\d{2})\s*[-–—]\s*(20\d{2}|present|current)",text,flags=re.I)
    total=0; now=datetime.date.today().year
    for a,b in ranges:
        total += max(0,(now if b.lower() in ("present","current") else int(b))-int(a))
    return min(total,40)

def has_any(text, terms):
    t=text.lower()
    return any(x.lower() in t for x in terms)

def score_candidate(cv_text,jd_text):
    cv=cv_text.lower(); jd=jd_text.lower(); score=100; deductions=[]; matches=[]
    reqs=[
        ("Meta Ads",["meta ads","facebook ads","instagram ads"],12,["meta ads","facebook ads","instagram ads"]),
        ("Google Ads",["google ads","google adwords"],12,["google ads","google adwords"]),
        ("Email Marketing",["email campaigns","email marketing","mailchimp"],10,["email marketing","mailchimp","hubspot","newsletter","email campaigns"]),
        ("GA4 / Google Analytics",["google analytics","ga4"],10,["ga4","google analytics"]),
        ("Canva",["canva"],6,["canva"]),
        ("Copywriting",["copy in nigerian tone","copywriting","write copy","social copy"],8,["copywriting","copywriter","social copy","ad copy","content copy"]),
    ]
    for label,jd_terms,pts,cv_terms in reqs:
        if has_any(jd,jd_terms):
            if has_any(cv,cv_terms):
                if label=="Meta Ads" and ("boosted posts only" in cv or "basic meta ads" in cv):
                    d=max(1,pts//2); score-=d; deductions.append(f"Limited {label}: -{d}pts")
                else: matches.append(label)
            else:
                score-=pts; deductions.append(f"No {label}: -{pts}pts")
    if has_any(jd,["fintech","banking"]):
        if has_any(cv,["kuda","opay","carbon","flutterwave","bank","fintech","gtbank","access bank","zenith"]):
            matches.append("Fintech/banking experience")
        else: score-=10; deductions.append("No fintech/banking experience: -10pts")
    yrs=years_experience(cv_text)
    if re.search(r"\b2\+\s*years|\b2\s+years",jd):
        if yrs>=2: matches.append("2+ years experience")
        else: score-=15; deductions.append("Less than 2 years relevant experience: -15pts")
    if "lagos" in jd and "hybrid" in jd and "lagos" not in cv:
        score-=5; deductions.append("Not Lagos-based for hybrid role: -5pts")
    return max(0,min(100,score)),matches,deductions,yrs

def ranking_group(score):
    return "Excellent" if score>=90 else "Good" if score>=70 else "Moderate" if score>=50 else "Maybe" if score>=30 else "Do Not Hire"

def verdict(score):
    return "Excellent match; shortlist immediately." if score>=90 else "Strong candidate; shortlist with minor gaps." if score>=70 else "Possible fit; review gaps carefully." if score>=50 else "Poor match; do not prioritize."

def make_charts(df):
    figs=[]; order=["Excellent","Good","Moderate","Maybe","Do Not Hire"]
    counts=df["Ranking Group"].value_counts().reindex(order,fill_value=0); nz=counts[counts>0]
    fig,ax=plt.subplots(figsize=(7,4.5)); ax.pie(nz.values,labels=nz.index,autopct="%1.0f%%",startangle=90); ax.set_title("Candidate Ranking Distribution"); figs.append(("Candidate Ranking Distribution",fig))
    fig,ax=plt.subplots(figsize=(7,4.5)); ax.hist(df["Fit %"],bins=min(10,max(3,int(df["Fit %"].nunique())))); ax.set_xlabel("Fit Score (%)"); ax.set_ylabel("Number of Candidates"); ax.set_title("Score Distribution"); figs.append(("Score Distribution",fig))
    top=df.sort_values("Fit %",ascending=False).head(10).iloc[::-1]
    fig,ax=plt.subplots(figsize=(7,4.8)); ax.barh(top["Name"],top["Fit %"]); ax.set_xlim(0,100); ax.set_xlabel("Fit Score (%)"); ax.set_title("Top 10 Candidates"); figs.append(("Top 10 Candidates",fig))
    return figs

def fig_png(fig):
    b=io.BytesIO(); fig.savefig(b,format="png",dpi=150,bbox_inches="tight"); plt.close(fig); b.seek(0); return b

def make_excel(df,details):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        pd.DataFrame({"Client Detail":list(details.keys())+["Total Candidates"],"Value":list(details.values())+[len(df)]}).to_excel(w,index=False,sheet_name="Summary")
        df.to_excel(w,index=False,sheet_name="Master Ranking")
        for g in ["Excellent","Good","Moderate","Maybe","Do Not Hire"]: df[df["Ranking Group"]==g].to_excel(w,index=False,sheet_name=g)
    return b.getvalue()

def make_docx(df,details):
    d=Document(); d.add_heading("CONFIDENTIAL CANDIDATE SCREENING REPORT",0)
    for k,v in details.items():
        if v: d.add_paragraph(f"{k}: {v}")
    d.add_paragraph(f"Screening date: {datetime.date.today().strftime('%d %B %Y')}")
    d.add_heading("Executive Summary",1); d.add_paragraph(f"{len(df)} candidates were screened, scored and ranked against the supplied Job Description.")
    for g in ["Excellent","Good","Moderate","Maybe","Do Not Hire"]:
        n=int((df["Ranking Group"]==g).sum())
        if n: d.add_paragraph(f"{g}: {n}",style="List Bullet")
    d.add_heading("Top Recommended Candidates",1)
    for _,r in df.sort_values("Fit %",ascending=False).head(10).iterrows(): d.add_paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})",style="List Number")
    for title,fig in make_charts(df): d.add_heading(title,1); d.add_picture(fig_png(fig),width=Inches(5.8))
    d.add_paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener.")
    b=io.BytesIO(); d.save(b); return b.getvalue()

def make_pdf(df,details):
    b=io.BytesIO(); styles=getSampleStyleSheet(); story=[Paragraph("CONFIDENTIAL CANDIDATE SCREENING REPORT",styles["Title"]),Spacer(1,12)]
    for k,v in details.items():
        if v: story.append(Paragraph(f"<b>{k}:</b> {v}",styles["BodyText"]))
    story += [Paragraph(f"<b>Screening date:</b> {datetime.date.today().strftime('%d %B %Y')}",styles["BodyText"]),Spacer(1,16),Paragraph("Executive Summary",styles["Heading1"]),Paragraph(f"{len(df)} candidates were screened, scored and ranked against the supplied Job Description.",styles["BodyText"]),Spacer(1,10)]
    rows=[["Ranking Group","Candidates"]]
    for g in ["Excellent","Good","Moderate","Maybe","Do Not Hire"]:
        n=int((df["Ranking Group"]==g).sum())
        if n: rows.append([g,n])
    t=Table(rows); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold")]))
    story += [t,Spacer(1,16),Paragraph("Top Recommended Candidates",styles["Heading1"])]
    for _,r in df.sort_values("Fit %",ascending=False).head(10).iterrows(): story.append(Paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})",styles["BodyText"]))
    for title,fig in make_charts(df): story += [Spacer(1,12),Paragraph(title,styles["Heading1"]),Image(fig_png(fig),width=430,height=260)]
    story += [Spacer(1,20),Paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener.",styles["Italic"])]
    SimpleDocTemplate(b,pagesize=A4,leftMargin=36,rightMargin=36,topMargin=40,bottomMargin=40).build(story); return b.getvalue()

st.title("Opportunity Hub Screener v2.4")
st.caption("Professional CV screening, batch processing and client-ready reporting")
with st.expander("Client & Recruitment Details",expanded=True):
    a,b=st.columns(2); client=a.text_input("Client / Company Name"); role=b.text_input("Recruitment Role / Job Title")
    a,b=st.columns(2); address=a.text_input("Client Address"); officer=b.text_input("Client Contact Officer")
    email=st.text_input("Client Contact Email Address")
details={"Prepared exclusively for":client,"Client Address":address,"Client Contact Officer":officer,"Client Contact Email":email,"Recruitment Project":role}

st.subheader("Job Description")
jd_mode=st.radio("JD source",["Paste JD","Upload JD File","Public JD Link"],horizontal=True); jd_text=""
if jd_mode=="Paste JD": jd_text=st.text_area("Paste Job Description",height=180)
elif jd_mode=="Upload JD File":
    f=st.file_uploader("Upload JD (PDF, DOCX or TXT)",type=["pdf","docx","txt"],key="jd")
    if f:
        jd_text=read_document(f.name,f.getvalue())
        st.success("JD file read successfully.") if jd_text.strip() else st.error("No readable text extracted. A scanned PDF may require OCR.")
else:
    u=st.text_input("Paste public JD file link")
    if u:
        try:
            n,data,ct=download_public_file(u); jd_text=read_document(n,data,ct); st.success("JD link read successfully.")
        except Exception as e: st.error(f"JD link error: {e}")

st.subheader("CV Intake")
mode=st.radio("Choose intake method",["Public CV Links","ZIP Batch","Individual CV Files"],horizontal=True); cv_items=[]; links=""
if mode=="Public CV Links": links=st.text_area("Paste one public CV link per line")
elif mode=="ZIP Batch":
    z=st.file_uploader("Upload one ZIP containing CVs",type=["zip"],key="zip")
    if z:
        try:
            with zipfile.ZipFile(io.BytesIO(z.getvalue())) as ar:
                for n in ar.namelist():
                    if not n.endswith("/"):
                        data=ar.read(n)
                        if detect_extension(n,data) in SUPPORTED: cv_items.append((Path(n).name,data,""))
            st.success(f"{len(cv_items)} supported CV files found.")
        except Exception as e: st.error(f"ZIP error: {e}")
else:
    up=st.file_uploader("Upload one or more CVs (PDF, DOCX or TXT)",type=["pdf","docx","txt"],accept_multiple_files=True,key="cvs")
    if up:
        cv_items=[(f.name,f.getvalue(),"") for f in up]; st.success(f"{len(cv_items)} CV file(s) ready.")

if st.button("Screen CVs",type="primary"):
    if not client or not role: st.error("Enter Client / Company Name and Recruitment Role / Job Title.")
    elif not jd_text.strip(): st.error("Provide a readable Job Description.")
    else:
        if mode=="Public CV Links":
            for u in [x.strip() for x in links.splitlines() if x.strip()]:
                try: cv_items.append(download_public_file(u))
                except Exception as e: st.warning(f"Could not download a CV: {e}")
        rows=[]; failures=[]; progress=st.progress(0)
        for i,(fn,data,ct) in enumerate(cv_items):
            try:
                text=read_document(fn,data,ct)
                if len(text.strip())<20: raise ValueError("No readable text extracted from this file.")
                score,matches,deductions,yrs=score_candidate(text,jd_text)
                rows.append({"Name":extract_name(text,fn),"Fit %":score,"2-Line Verdict":verdict(score),"Why Not 100%":"; ".join(deductions) if deductions else "No material requirement missed.","Red Flag":"; ".join(deductions[:2]) if deductions else "None material","Green Flag":"; ".join(matches[:3]) if matches else "Limited evidence","Years Exp":yrs,"Skills Match":", ".join(matches) if matches else "None detected","Resume Link":fn,"Ranking Group":ranking_group(score)})
            except Exception as e: failures.append({"File":fn,"Error":str(e)})
            progress.progress((i+1)/max(len(cv_items),1))
        if rows:
            df=pd.DataFrame(rows).sort_values("Fit %",ascending=False).reset_index(drop=True)
            st.session_state.results=df; st.session_state.failures=pd.DataFrame(failures)
            st.session_state.xlsx=make_excel(df,details); st.session_state.csv=df.to_csv(index=False).encode(); st.session_state.docx=make_docx(df,details); st.session_state.pdf=make_pdf(df,details)
            st.success(f"Screening completed: {len(df)} CVs processed successfully.")
        else: st.error("No CV was successfully processed.")

if "results" in st.session_state:
    df=st.session_state.results; st.subheader("Screening Results")
    st.dataframe(df[["Name","Fit %","Ranking Group","2-Line Verdict","Why Not 100%","Years Exp","Skills Match"]],use_container_width=True)
    st.subheader("Reports and Charts")
    for title,fig in make_charts(df): st.pyplot(fig)
    st.download_button("Download Master Workbook",st.session_state.xlsx,"candidate_screening_workbook.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download CSV",st.session_state.csv,"candidate_screening_results.csv","text/csv")
    st.download_button("Download Professional DOCX Report",st.session_state.docx,"candidate_screening_report.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.download_button("Download Professional PDF Report",st.session_state.pdf,"candidate_screening_report.pdf","application/pdf")
if "failures" in st.session_state and not st.session_state.failures.empty:
    with st.expander("Files that could not be processed"): st.dataframe(st.session_state.failures,use_container_width=True)
st.divider(); st.subheader("Enterprise Scale Foundation")
st.info("For thousands or millions of CVs, Streamlit remains the client control panel while intake moves to durable object storage, batch manifests, queues, parallel workers, checkpoints, retries and a results database.")
