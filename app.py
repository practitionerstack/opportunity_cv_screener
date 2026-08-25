import streamlit as st
import pandas as pd, io, re, zipfile, hashlib
from pathlib import Path
from datetime import date
import fitz
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="Opportunity Hub CV Screener v2.7.6", page_icon="🎯", layout="wide")
PREF=["preferably","preferred","nice to have","nice-to-have","advantage","bonus","plus","desirable"]
REQ=["must have","required","essential","mandatory","minimum requirement"]
SKILLS={
"Meta Ads":[r"\\bmeta ads?\\b",r"\\bfacebook ads?\\b",r"\\binstagram ads?\\b"],
"Google Ads":[r"\\bgoogle ads?\\b",r"\\badwords\\b"],
"Copywriting":[r"\\bcopywriting\\b",r"\\bwrite copy\\b",r"\\bad copy\\b",r"\\bsocial copy\\b"],
"GA4 / Google Analytics":[r"\\bga4\\b",r"\\bgoogle analytics\\b"],
"Canva":[r"\\bcanva\\b"],"HubSpot":[r"\\bhubspot\\b"],
"Basic Design":[r"\\bbasic design\\b",r"\\bgraphic design\\b"],
"TikTok Ads":[r"\\btiktok ads?\\b"],
"Fintech / banking experience":[r"\\bfintech\\b",r"\\bbanking\\b",r"\\bbank\\b",r"\\bkuda\\b",r"\\bopay\\b",r"\\bcarbon\\b",r"\\bflutterwave\\b",r"\\bgtbank\\b"]}

def extract(name,data):
 e=Path(name).suffix.lower()
 try:
  if e==".pdf":
   with fitz.open(stream=data,filetype="pdf") as d: t="\\n".join(p.get_text() for p in d)
   if len(t.strip())>20:return t,""
   r=PdfReader(io.BytesIO(data));return "\\n".join(p.extract_text() or "" for p in r.pages),""
  if e==".docx":
   d=Document(io.BytesIO(data));return "\\n".join(p.text for p in d.paragraphs),""
  if e==".txt":return data.decode(errors="ignore"),""
 except Exception as x:return "",str(x)
 return "",f"Unsupported {e}"

def lines(t):
 return [x.strip(" -•\\t") for x in re.split(r"[\\r\\n]+",t) if x.strip()]

def cat(s,section):
 z=(s+" "+section).lower()
 if any(q in z for q in PREF):return "Preferred","Preferred qualifier detected"
 if any(q in z for q in REQ):return "Required","Mandatory qualifier/section detected"
 if section.lower() in ["nice to have","preferred","desirable","advantage"]:return "Preferred","Preferred section"
 return "Required","Default required criterion"

def parse_jd(t):
 section="General"; out=[]; seen=set()
 for line in lines(t):
  m=re.match(r"^(Must Have|Nice to Have|Required|Essential|Preferred|Desirable|Requirements?)\\s*:\\s*(.*)$",line,re.I)
  if m: section=m.group(1); line=m.group(2).strip() or ""; 
  if not line:continue
  c,reason=cat(line,section)
  for name,pats in SKILLS.items():
   if any(re.search(p,line,re.I) for p in pats):
    if name not in seen:
     out.append({"name":name,"category":c,"source":section,"source_text":line,"type":"preference" if name=="Fintech / banking experience" else "skill","weight":4 if c=="Required" else 1,"reason":reason});seen.add(name)
  y=re.search(r"(\\d+)\\s*\\+?\\s*years?",line,re.I)
  if y and any(k in line.lower() for k in ["marketing","experience","digital"]):
   name=f"{y.group(1)}+ years relevant experience"
   if name not in seen:out.append({"name":name,"category":c,"source":section,"source_text":line,"type":"experience","years":int(y.group(1)),"weight":4 if c=="Required" else 1,"reason":reason});seen.add(name)
  money=re.search(r"(₦|ngn)\\s*([\\d,.]+)\\s*(m|million|k|thousand)?",line,re.I)
  if money and any(k in line.lower() for k in ["budget","spend","campaign"]):
   n=float(money.group(2).replace(",",""));u=(money.group(3) or "").lower()
   n*=1000000 if u in ["m","million"] else 1000 if u in ["k","thousand"] else 1
   name=f"Monthly ad budget ≥ ₦{n:,.0f}"
   if name not in seen:out.append({"name":name,"category":c,"source":section,"source_text":line,"type":"threshold","threshold":n,"weight":4 if c=="Required" else 1,"reason":reason});seen.add(name)
 return out

def yrs(t):
 z=0
 for a,b in re.findall(r"\\b(20\\d{2})\\s*(?:-|–|to)\\s*(20\\d{2}|present|current)\\b",t.lower()):
  z+=max(0,(date.today().year if b in ["present","current"] else int(b))-int(a))
 return z

def skill_ev(t,pats):
 for l in lines(t):
  if any(re.search(p,l,re.I) for p in pats):
   if re.search(r"\\b(no|without|lack of|lacks?)\\b",l,re.I):return 0,"Negative evidence","Explicitly absent",l
   low=l.lower()
   if any(x in low for x in ["basic","assisted","support","boosted posts","sometimes"]):return .35,"Positive evidence","Weak / limited evidence",l
   if any(x in low for x in ["managed","led","budget","certified","proficient","optimized","built"]):return 1,"Positive evidence","Direct evidence",l
   return .7,"Positive evidence","Evidence present",l
 return 0,"No evidence","Not demonstrated in CV",""

def budget_ev(t,threshold):
 neg=None;best=None
 for l in lines(t):
  if re.search(r"\\bno\\b.*\\bbudget\\b",l,re.I):neg=l
  for m in re.finditer(r"(₦|ngn)\\s*([\\d,.]+)\\s*(m|million|k|thousand)?",l,re.I):
   n=float(m.group(2).replace(",",""));u=(m.group(3) or "").lower();n*=1000000 if u in ["m","million"] else 1000 if u in ["k","thousand"] else 1
   if any(k in l.lower() for k in ["budget","ads","campaign"]):best=max(best or (0,""),(n,l))
 if best:return (1 if best[0]>=threshold else 0),"Positive evidence","Meets threshold" if best[0]>=threshold else "Below threshold",best[1]
 if neg:return 0,"Negative evidence","Explicitly absent",neg
 return 0,"No evidence","Not demonstrated in CV",""

def score(t,reqs):
 rows=[];earned=0;poss=0;pe=0;pp=0;missing=[]
 for r in reqs:
  if r["type"]=="experience":
   y=yrs(t);level=min(1,y/r["years"]);et="Positive evidence";status="Meets requirement" if level==1 else "Below requirement";ev=f"Estimated dated experience: {y} year(s)"
  elif r["type"]=="threshold":level,et,status,ev=budget_ev(t,r["threshold"])
  else:level,et,status,ev=skill_ev(t,SKILLS[r["name"]])
  pts=r["weight"]*level
  if r["category"]=="Required":earned+=pts;poss+=r["weight"];missing += [r["name"]] if level==0 else []
  else:pe+=pts;pp+=r["weight"]
  rows.append({"Requirement":r["name"],"Category":"Must-have" if r["category"]=="Required" else "Preferred","Source / Provenance":r["source"],"JD Source Excerpt":r["source_text"],"Requirement Type":r["type"],"Evidence Type":et,"Evidence Level":level,"Status":status,"Evidence":ev,"Points Earned":round(pts,2)})
 mandatory=100*earned/(poss or 1);bonus=5*pe/pp if pp else 0;final=min(100,mandatory+bonus)
 return round(final),round(mandatory,1),round(bonus,1),missing,rows

st.title("🎯 Opportunity Hub CV Screener v2.7.6")
st.caption("Compound Requirement & Threshold Extraction • v2.7.5 preserved as rollback baseline")
jd_text=st.text_area("Paste Job Description",height=160)
jd_file=st.file_uploader("Upload JD",type=["pdf","docx","txt"])
cvs=st.file_uploader("Upload CVs or ZIP batch",type=["pdf","docx","txt","zip"],accept_multiple_files=True)
if "r" not in st.session_state:st.session_state.r=None
if st.button("Screen CVs"):
 jd=jd_text
 if jd_file:
  x,e=extract(jd_file.name,jd_file.getvalue())
  if e:st.error(e);st.stop()
  jd+="\\n"+x
 reqs=parse_jd(jd)
 if not reqs:st.error("No validated requirements extracted. Screening stopped safely.");st.stop()
 results=[];audits={}
 for f in cvs or []:
  items=[]
  if f.name.lower().endswith(".zip"):
   with zipfile.ZipFile(io.BytesIO(f.getvalue())) as z:items=[(Path(i.filename).name,z.read(i)) for i in z.infolist() if not i.is_dir()]
  else:items=[(f.name,f.getvalue())]
  for fn,data in items:
   t,e=extract(fn,data)
   if e or not t:continue
   name=Path(fn).stem.replace("CV_","").replace("-"," ").replace("_"," ").strip().title()
   s,m,b,miss,a=score(t,reqs);g="Excellent" if s>=90 else "Good" if s>=70 else "Moderate" if s>=50 else "Do Not Hire"
   results.append({"Name":name,"File":fn,"Fit %":s,"Ranking Group":g,"Mandatory Score":m,"Preferred Bonus":b,"Years Exp":yrs(t),"Why Not 100%":"; ".join(miss)});audits[name]=a
 st.session_state.r=(reqs,pd.DataFrame(results).sort_values(["Fit %","Mandatory Score"],ascending=False),audits)
if st.session_state.r:
 reqs,df,audits=st.session_state.r
 st.subheader("JD Requirement Lock");lock=pd.DataFrame(reqs);st.dataframe(lock,use_container_width=True)
 st.subheader("Master Ranking");st.dataframe(df,use_container_width=True)
 st.subheader("Candidate-by-Candidate Scoring Audit")
 for n in df["Name"]:
  with st.expander(n):st.dataframe(pd.DataFrame(audits[n]),use_container_width=True)
 out=io.BytesIO()
 with pd.ExcelWriter(out,engine="openpyxl") as w:
  df.to_excel(w,index=False,sheet_name="Master Ranking");lock.to_excel(w,index=False,sheet_name="JD Requirements")
  pd.concat([pd.DataFrame([{"Candidate":n,**x} for x in rows]) for n,rows in audits.items()],ignore_index=True).to_excel(w,index=False,sheet_name="Scoring Audit")
 st.download_button("Download Workbook",out.getvalue(),"opportunity_hub_v2_7_6.xlsx")
 st.download_button("Download CSV",df.to_csv(index=False).encode(),"opportunity_hub_v2_7_6.csv","text/csv")
 st.caption("Screened, sorted, scored and ranked by Opportunity Hub Screener.")
