
import io, re, zipfile, datetime
from pathlib import Path

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

st.set_page_config(page_title="Opportunity Hub Screener v2.3", layout="wide")

def read_document(filename, data):
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    if ext == ".docx":
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)
    if ext in [".txt", ".csv"]:
        return data.decode("utf-8", errors="ignore")
    return ""

def google_drive_download_url(url):
    match = re.search(r"/d/([^/]+)", url)
    if "drive.google.com" in url and match:
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
    return url

def download_public_file(url, timeout=120):
    r = requests.get(google_drive_download_url(url), timeout=timeout)
    r.raise_for_status()
    return r.content

def candidate_name(text, fallback):
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    for line in lines[:10]:
        words = line.split()
        if 2 <= len(words) <= 5 and len(line) < 70:
            bad = ["curriculum", "resume", "cv", "experience", "skills", "education"]
            if not any(x in line.lower() for x in bad):
                return line.title()
    return Path(fallback).stem.replace("_", " ").replace("-", " ").title()

def years_experience(text):
    years = sorted(set(int(x) for x in re.findall(r"\b20\d{2}\b", text)))
    if len(years) >= 2:
        return max(years) - min(years)
    return 0

def contains_any(text, terms):
    return any(term.lower() in text.lower() for term in terms)

def score_candidate(cv_text, jd_text):
    cv = cv_text.lower()
    jd = jd_text.lower()
    checks = [
        ("Meta Ads", ["meta ads", "facebook ads", "instagram ads"], 15),
        ("Google Ads", ["google ads", "google adwords"], 15),
        ("GA4 / Google Analytics", ["ga4", "google analytics"], 10),
        ("Email Marketing", ["email marketing", "mailchimp", "hubspot", "newsletter"], 10),
        ("Canva", ["canva"], 5),
        ("Copywriting", ["copywriting", "copywriter", "social copy", "ad copy"], 10),
    ]

    evidence, total, possible = [], 0, 0

    for label, terms, weight in checks:
        if contains_any(jd, terms):
            possible += weight
            hit = contains_any(cv, terms)
            explicit_negative = f"no {label.lower()}" in cv
            basic = ("basic" in cv and any(t in cv for t in terms)) or ("boosted posts" in cv and label == "Meta Ads")
            points = 0 if explicit_negative or not hit else (round(weight * 0.4) if basic else weight)
            total += points
            evidence.append((label, points, weight))

    if contains_any(jd, ["fintech", "banking", "bank"]):
        possible += 10
        fintech_terms = ["fintech", "bank", "kuda", "opay", "carbon", "flutterwave", "gtbank", "access bank", "zenith"]
        points = 10 if contains_any(cv, fintech_terms) else 0
        total += points
        evidence.append(("Fintech / Banking Experience", points, 10))

    if "lagos" in jd:
        possible += 5
        points = 5 if "lagos" in cv else 0
        total += points
        evidence.append(("Lagos Location", points, 5))

    yrs = years_experience(cv_text)
    if re.search(r"\b2\+?\s*years", jd):
        possible += 10
        points = 10 if yrs >= 2 else 0
        total += points
        evidence.append(("2+ Years Relevant Experience", points, 10))

    score = round(100 * total / max(possible, 1))
    matched = [name for name, pts, _ in evidence if pts > 0]
    gaps = [name for name, pts, _ in evidence if pts == 0]
    partial = [name for name, pts, mx in evidence if 0 < pts < mx]
    return min(score, 100), evidence, matched, gaps, partial, yrs

def ranking_group(score):
    if score >= 90: return "Excellent"
    if score >= 70: return "Good"
    if score >= 50: return "Moderate"
    if score >= 30: return "Maybe"
    return "Do Not Hire"

def verdict(score):
    if score >= 90: return "Excellent match; shortlist immediately."
    if score >= 70: return "Strong candidate; shortlist with minor gaps."
    if score >= 50: return "Possible fit; review gaps carefully."
    return "Poor match; do not prioritize."

def make_charts(df):
    charts = []
    order = ["Excellent", "Good", "Moderate", "Maybe", "Do Not Hire"]
    counts = df["Ranking Group"].value_counts().reindex(order, fill_value=0)

    fig, ax = plt.subplots()
    ax.pie(counts.values, labels=counts.index, autopct="%1.0f%%")
    ax.set_title("Candidate Ranking Distribution")
    charts.append(("Candidate Ranking Distribution", fig))

    fig, ax = plt.subplots()
    ax.hist(df["Fit %"], bins=10)
    ax.set_xlabel("Fit Score")
    ax.set_ylabel("Number of Candidates")
    ax.set_title("Score Distribution")
    charts.append(("Score Distribution", fig))

    top = df.sort_values("Fit %", ascending=False).head(10).iloc[::-1]
    fig, ax = plt.subplots()
    ax.barh(top["Name"], top["Fit %"])
    ax.set_xlabel("Fit %")
    ax.set_title("Top 10 Candidates")
    charts.append(("Top Candidates", fig))
    return charts

def fig_to_png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

def excel_report(df, client, role):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary = pd.DataFrame({
            "Metric": ["Client", "Recruitment Project", "Total Candidates"],
            "Value": [client, role, len(df)]
        })
        summary.to_excel(writer, index=False, sheet_name="Summary")
        df.to_excel(writer, index=False, sheet_name="Master Ranking")
        for group in ["Excellent", "Good", "Moderate", "Maybe", "Do Not Hire"]:
            df[df["Ranking Group"] == group].to_excel(writer, index=False, sheet_name=group)
    return buf.getvalue()

def docx_report(df, client, role):
    doc = Document()
    doc.add_heading("CONFIDENTIAL CANDIDATE SCREENING REPORT", 0)
    doc.add_paragraph(f"Prepared exclusively for: {client}")
    doc.add_paragraph(f"Recruitment project: {role}")
    doc.add_paragraph(f"Screening date: {datetime.date.today().strftime('%d %B %Y')}")
    doc.add_heading("Executive Summary", 1)
    doc.add_paragraph(f"{len(df)} candidate records were screened, scored and ranked against the stated requirements.")
    for group, count in df["Ranking Group"].value_counts().items():
        doc.add_paragraph(f"{group}: {count}", style="List Bullet")
    doc.add_heading("Top Recommended Candidates", 1)
    for _, row in df.sort_values("Fit %", ascending=False).head(10).iterrows():
        doc.add_paragraph(f"{row['Name']} — {row['Fit %']}% ({row['Ranking Group']})", style="List Number")
    for title, fig in make_charts(df):
        doc.add_heading(title, 1)
        doc.add_picture(fig_to_png(fig), width=Inches(5.8))
    doc.add_paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener.")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def pdf_report(df, client, role):
    buf = io.BytesIO()
    styles = getSampleStyleSheet()
    story = [
        Paragraph("CONFIDENTIAL CANDIDATE SCREENING REPORT", styles["Title"]),
        Spacer(1, 12),
        Paragraph(f"<b>Prepared exclusively for:</b> {client}", styles["BodyText"]),
        Paragraph(f"<b>Recruitment project:</b> {role}", styles["BodyText"]),
        Paragraph(f"<b>Screening date:</b> {datetime.date.today().strftime('%d %B %Y')}", styles["BodyText"]),
        Spacer(1, 16),
        Paragraph("Executive Summary", styles["Heading1"]),
        Paragraph(f"{len(df)} candidates were screened, scored and ranked.", styles["BodyText"]),
        Spacer(1, 10)
    ]
    summary = [["Ranking Group", "Candidates"]]
    for group, count in df["Ranking Group"].value_counts().items():
        summary.append([group, int(count)])
    table = Table(summary)
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")
    ]))
    story += [table, Spacer(1, 16), Paragraph("Top Recommended Candidates", styles["Heading1"])]
    for _, row in df.sort_values("Fit %", ascending=False).head(10).iterrows():
        story.append(Paragraph(f"{row['Name']} — {row['Fit %']}% ({row['Ranking Group']})", styles["BodyText"]))
    for title, fig in make_charts(df):
        story += [Spacer(1, 12), Paragraph(title, styles["Heading1"])]
        story.append(Image(fig_to_png(fig), width=430, height=260))
    story += [Spacer(1, 20), Paragraph("Screened, sorted, scored and ranked by Opportunity Hub Screener.", styles["Italic"])]
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=40, bottomMargin=40).build(story)
    return buf.getvalue()

st.title("Opportunity Hub Screener v2.3")
st.caption("Professional screening, client reporting and enterprise-scale intake foundation")

client = st.text_input("Client / Company Name")
role = st.text_input("Recruitment Project / Role")
jd = st.text_area("Paste Job Description", height=180)

st.subheader("CV Intake")
mode = st.radio("Choose intake method", [
    "Public file links (mobile-safe)",
    "ZIP batch",
    "Individual upload"
])

files = []
if mode == "Public file links (mobile-safe)":
    links = st.text_area("Paste one public CV link per line")
elif mode == "ZIP batch":
    uploaded_zip = st.file_uploader("Upload ONE ZIP containing CV files", type=["zip"])
    links = ""
    if uploaded_zip:
        try:
            with zipfile.ZipFile(io.BytesIO(uploaded_zip.getvalue())) as archive:
                files = [(Path(n).name, archive.read(n)) for n in archive.namelist()
                         if n.lower().endswith((".pdf", ".docx", ".txt")) and not n.endswith("/")]
            st.success(f"{len(files)} supported CV files found.")
        except Exception as exc:
            st.error(f"ZIP could not be read: {exc}")
else:
    uploaded = st.file_uploader("Upload CV files", type=["pdf", "docx", "txt"], accept_multiple_files=True)
    links = ""
    if uploaded:
        files = [(f.name, f.getvalue()) for f in uploaded]

if st.button("Screen Candidates", type="primary"):
    if not client or not role or not jd:
        st.error("Enter the client name, recruitment project and Job Description.")
    else:
        if mode == "Public file links (mobile-safe)":
            for url in [x.strip() for x in links.splitlines() if x.strip()]:
                try:
                    files.append((url.split("/")[-1].split("?")[0] or "CV.pdf", download_public_file(url)))
                except Exception as exc:
                    st.warning(f"Could not download one CV: {exc}")

        if not files:
            st.error("No CVs are available for screening.")
        else:
            rows = []
            progress = st.progress(0)
            for i, (filename, data) in enumerate(files):
                try:
                    text = read_document(filename, data)
                    score, evidence, matched, gaps, partial, yrs = score_candidate(text, jd)
                    rows.append({
                        "Name": candidate_name(text, filename),
                        "Fit %": score,
                        "2-Line Verdict": verdict(score),
                        "Why Not 100%": "; ".join(gaps[:4]) or "No material gaps detected.",
                        "Red Flag": "; ".join(gaps[:2]) or "None material",
                        "Green Flag": "; ".join(matched[:3]) or "Limited evidence",
                        "Years Exp": yrs,
                        "Skills Match": ", ".join(matched) or "None detected",
                        "Resume Link": filename,
                        "Ranking Group": ranking_group(score),
                        "Audit Trail": "; ".join(f"{name}: {pts}/{maximum}" for name, pts, maximum in evidence)
                    })
                except Exception as exc:
                    st.warning(f"Failed to process {filename}: {exc}")
                progress.progress((i + 1) / len(files))

            if rows:
                results = pd.DataFrame(rows).sort_values("Fit %", ascending=False).reset_index(drop=True)
                st.dataframe(results[["Name", "Fit %", "Ranking Group", "2-Line Verdict", "Why Not 100%"]], use_container_width=True)

                for title, fig in make_charts(results):
                    st.subheader(title)
                    st.pyplot(fig)

                st.download_button("Download XLSX Workbook", excel_report(results, client, role),
                                   f"{client}_screening.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.download_button("Download CSV", results.to_csv(index=False).encode("utf-8"),
                                   f"{client}_screening.csv", "text/csv")
                st.download_button("Download Professional DOCX Report", docx_report(results, client, role),
                                   f"{client}_screening_report.docx",
                                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.download_button("Download Professional PDF Report", pdf_report(results, client, role),
                                   f"{client}_screening_report.pdf", "application/pdf")

st.divider()
st.subheader("Enterprise Scale Foundation")
st.info("For very large jobs, the screening logic remains the same while intake moves to object storage, a batch manifest, queue and parallel workers. This prevents million-CV contracts from requiring a rewrite of the scoring engine.")
