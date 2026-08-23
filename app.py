import streamlit as st, pandas as pd, re, io, zipfile, hashlib, requests
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, parse_qs
from pypdf import PdfReader
from docx import Document
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

st.set_page_config(page_title='Opportunity Hub CV Screener v2.5', layout='wide')
SUPPORTED={'.pdf','.docx','.txt','.csv','.xlsx','.xls'}

def clean(x): return re.sub(r'\s+',' ',str(x or '')).strip()
def h(data): return hashlib.sha256(data).hexdigest()

def drive_url(url):
    m=re.search(r'/d/([^/]+)',url)
    if m:return f'https://drive.google.com/uc?export=download&id={m.group(1)}'
    q=parse_qs(urlparse(url).query)
    return f'https://drive.google.com/uc?export=download&id={q["id"][0]}' if q.get('id') else url

def download(url):
    r=requests.get(drive_url(url),timeout=60,allow_redirects=True,headers={'User-Agent':'Mozilla/5.0'});r.raise_for_status()
    if 'text/html' in (r.headers.get('content-type') or '').lower() and not r.content.startswith(b'%PDF'):
        raise ValueError('Link returned a webpage instead of a downloadable document.')
    return r.content

def extract(data,name):
    ext=Path(name).suffix.lower()
    if data.startswith(b'%PDF') or ext=='.pdf':
        return clean('\n'.join(p.extract_text() or '' for p in PdfReader(io.BytesIO(data)).pages))
    if ext=='.docx':
        d=Document(io.BytesIO(data));return clean('\n'.join(p.text for p in d.paragraphs))
    if ext=='.txt': return clean(data.decode('utf-8',errors='ignore'))
    if ext=='.csv': return clean(pd.read_csv(io.BytesIO(data)).to_string(index=False))
    if ext in {'.xlsx','.xls'}:
        return clean('\n'.join(df.to_string(index=False) for df in pd.read_excel(io.BytesIO(data),sheet_name=None).values()))
    if data[:2]==b'PK':
        try:return clean('\n'.join(p.text for p in Document(io.BytesIO(data)).paragraphs))
        except: pass
    raise ValueError(f'Unsupported or unreadable file: {name}')

def parse_links(s): return [x.strip() for x in re.split(r'[\n,]+',s or '') if x.strip()]
def name_from(text,filename):
    for line in text.splitlines()[:12]:
        s=clean(line)
        if 3<=len(s)<=60 and 2<=len(s.split())<=4 and not re.search(r'cv|resume|experience|email|phone|address',s,re.I): return s.title() if s.isupper() else s
    return clean(re.sub(r'(?i)^cv[_\-\s]*\d+[_\-\s]*','',Path(filename).stem).replace('_',' ').replace('-',' '))
def years(text):
    total=0;now=datetime.now().year
    for a,b in re.findall(r'(20\d{2})\s*[-–]\s*(20\d{2}|present|current)',text,re.I):
        total=max(total,(now if b.lower() in {'present','current'} else int(b))-int(a))
    return total
def location(text):
    for x in ['lagos','abuja','ibadan','port harcourt','enugu','remote']:
        if x in text.lower(): return x.title()
    return 'Not evidenced'

def requirements(jd):
    j=jd.lower();R=[]
    def add(label,terms,w,required=True,partial=[]):R.append((label,terms,w,required,partial))
    m=re.search(r'(\d+)\+?\s*years?',j)
    if m:add(f'{m.group(1)}+ years relevant experience',[],15,True)
    if 'meta ads' in j or 'facebook ads' in j:add('Meta Ads',['meta ads','facebook ads'],15,True,['boosted posts','boosting posts'])
    if 'google ads' in j:add('Google Ads',['google ads','adwords'],15,True)
    if 'email' in j:add('Email marketing',['email marketing','mailchimp','hubspot','newsletter','email campaign'],12,True)
    if 'ga4' in j or 'google analytics' in j:add('GA4 / Google Analytics',['ga4','google analytics'],10,True)
    if 'canva' in j:add('Canva',['canva'],6,True)
    if 'copy' in j or 'nigerian tone' in j:add('Copywriting / local tone',['copywriting','social copy','ad copy','content copy','wrote copy'],10,True)
    if 'fintech' in j or 'banking' in j:add('Fintech / banking experience',['kuda','opay','carbon','flutterwave','bank','fintech','payment'],7,False)
    if 'hubspot' in j:add('HubSpot',['hubspot'],3,False)
    if 'design' in j:add('Basic design',['canva','photoshop','illustrator','figma'],2,False)
    return R or [('Relevant professional experience',['experience'],50,True,[]),('Relevant skills',['skills'],50,True,[])]

def score(text,filename,jd):
    ev=[];total=sum(x[2] for x in requirements(jd));pts=0;yrs=years(text);tl=text.lower()
    for label,terms,w,required,partial in requirements(jd):
        match='No';e='Not explicitly evidenced';a=0
        if 'years relevant experience' in label:
            n=int(re.search(r'(\d+)\+',label).group(1))
            if yrs>=n:match,e,a='Yes',f'{yrs} years evidenced from CV dates',w
            elif yrs>0:match,e,a='Partial',f'{yrs} years evidenced; below requirement',w*.4
            else:e='No reliable dated experience evidence'
        else:
            hits=[t for t in terms if t in tl];ph=[t for t in partial if t in tl]
            if hits:match,e,a='Yes',', '.join(hits[:3]),w
            elif ph:match,e,a='Partial',', '.join(ph[:2]),w*.25
        pts+=a;ev.append({'Requirement':label,'Weight':w,'Match':match,'Evidence':e,'Awarded':round(a,1),'Required':required})
    fit=round(pts/total*100) if total else 0
    if any(x['Required'] and x['Match']!='Yes' for x in ev):fit=min(fit,95)
    group='Excellent' if fit>=90 else 'Good' if fit>=70 else 'Moderate' if fit>=50 else 'Do Not Hire'
    verdict={'Excellent':'Excellent match; shortlist immediately.','Good':'Strong candidate; shortlist with review.','Moderate':'Possible fit; gaps require careful review.','Do Not Hire':'Poor match; do not prioritize.'}[group]
    gaps=[f"{x['Requirement']} ({x['Weight']} pts)" for x in ev if x['Match']=='No']+[f"{x['Requirement']} partial ({x['Weight']-x['Awarded']:.0f} pts)" for x in ev if x['Match']=='Partial']
    return {'Name':name_from(text,filename),'Fit %':int(fit),'Ranking Group':group,'2-Line Verdict':verdict,'Why Not 100%':'; '.join(gaps[:4]) or 'All scored requirements explicitly evidenced','Years Exp':yrs,'Location':location(text),'Evidence Matrix':ev}

def charts(df,kind):
    fig,ax=plt.subplots(figsize=(7,4.2))
    if kind=='pie':
        c=df['Ranking Group'].value_counts();ax.pie(c.values,labels=c.index,autopct='%1.0f%%');ax.set_title('Candidate Ranking Distribution')
    elif kind=='hist':ax.hist(df['Fit %'],bins=min(10,max(3,len(df))),edgecolor='black');ax.set_xlabel('Fit Score (%)');ax.set_ylabel('Number of Candidates');ax.set_title('Score Distribution')
    else:
        d=df.head(10).iloc[::-1];ax.barh(d['Name'],d['Fit %']);ax.set_xlabel('Fit Score (%)');ax.set_title('Top 10 Candidates')
    fig.tight_layout();b=io.BytesIO();fig.savefig(b,format='png',dpi=160);plt.close(fig);b.seek(0);return b

def excel(df,meta):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine='openpyxl') as w:
        df.to_excel(w,sheet_name='Master Ranking',index=False)
        for g in ['Excellent','Good','Moderate','Do Not Hire']:df[df['Ranking Group']==g].to_excel(w,sheet_name=g,index=False)
        pd.DataFrame(meta.items(),columns=['Field','Value']).to_excel(w,sheet_name='Screening Summary',index=False)
    return b.getvalue()

def docx(df,meta):
    d=Document();d.add_heading('CONFIDENTIAL CANDIDATE SCREENING REPORT',0)
    for k,v in meta.items():d.add_paragraph(f'{k}: {v}')
    d.add_heading('Executive Summary',1);d.add_paragraph(f'{len(df)} candidates were screened, scored and ranked against the supplied Job Description.')
    d.add_heading('Top Recommended Candidates',1)
    for _,r in df.head(10).iterrows():d.add_paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})")
    import tempfile,os
    for title,k in [('Candidate Ranking Distribution','pie'),('Score Distribution','hist'),('Top 10 Candidates','top')]:
        d.add_page_break();d.add_heading(title,1);f=tempfile.NamedTemporaryFile(delete=False,suffix='.png');f.write(charts(df,k).getvalue());f.close();d.add_picture(f.name,width=6.2*inch);os.unlink(f.name)
    d.add_paragraph('Screened, sorted, scored and ranked by Opportunity Hub Screener.');b=io.BytesIO();d.save(b);return b.getvalue()

def pdf(df,meta):
    b=io.BytesIO();S=getSampleStyleSheet();story=[Paragraph('CONFIDENTIAL CANDIDATE SCREENING REPORT',S['Title']),Spacer(1,12)]
    for k,v in meta.items():story.append(Paragraph(f'<b>{k}:</b> {v}',S['BodyText']))
    story += [Spacer(1,12),Paragraph('Executive Summary',S['Heading1']),Paragraph(f'{len(df)} candidates were screened, scored and ranked against the supplied Job Description.',S['BodyText'])]
    c=df['Ranking Group'].value_counts().reset_index();c.columns=['Ranking Group','Candidates'];t=Table([c.columns.tolist()]+c.values.tolist());t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.5,colors.grey),('BACKGROUND',(0,0),(-1,0),colors.lightgrey)]));story += [Spacer(1,8),t,Spacer(1,14),Paragraph('Top Recommended Candidates',S['Heading1'])]
    for _,r in df.head(10).iterrows():story.append(Paragraph(f"{r['Name']} — {r['Fit %']}% ({r['Ranking Group']})",S['BodyText']))
    for title,k in [('Candidate Ranking Distribution','pie'),('Score Distribution','hist'),('Top 10 Candidates','top')]:story += [PageBreak(),Paragraph(title,S['Heading1']),RLImage(charts(df,k),width=6.2*inch,height=3.8*inch)]
    story += [Spacer(1,12),Paragraph('<i>Screened, sorted, scored and ranked by Opportunity Hub Screener.</i>',S['BodyText'])]
    SimpleDocTemplate(b,pagesize=A4,rightMargin=42,leftMargin=42,topMargin=42,bottomMargin=42).build(story);return b.getvalue()

st.title('🎯 Opportunity Hub — CV Screening Service v2.5');st.caption('Evidence-based screening • multi-format intake • duplicate detection • persistent results')
if 'result' not in st.session_state:st.session_state.result=None
with st.expander('Client & Recruitment Details',expanded=True):
    a,b=st.columns(2)
    with a:cn=st.text_input('Client Name');ca=st.text_input('Client Address');co=st.text_input('Client Contact Officer')
    with b:ce=st.text_input('Client Contact Email');role=st.text_input('Recruitment Role / Job');sd=st.date_input('Screening Date',datetime.now().date())
st.subheader('Job Description Intake');jdp=st.text_area('Paste Job Description (optional)',height=160);jdl=st.text_input('JD public/direct link (optional)');jdu=st.file_uploader('Upload Job Description — PDF, DOCX, TXT, CSV, XLSX',type=['pdf','docx','txt','csv','xlsx','xls'])
st.subheader('CV Intake');links=st.text_area('Paste one or more CV links (one per line)',height=110);files=st.file_uploader('Upload individual CV files',type=['pdf','docx','txt','csv','xlsx','xls'],accept_multiple_files=True);zup=st.file_uploader('Upload ZIP containing CVs',type=['zip']);st.info('Duplicate protection detects identical files and duplicate document content before scoring.')
if st.button('Screen CVs',type='primary'):
    errs=[];jd=''
    try:
        if jdu:jd=extract(jdu.getvalue(),jdu.name)
        elif jdl.strip():
            x=download(jdl.strip());jd=extract(x,'job.pdf' if x.startswith(b'%PDF') else 'job.docx')
        else:jd=jdp
        jd=clean(jd)
        if not jd:raise ValueError('No readable Job Description supplied.')
    except Exception as e:errs.append(f'JD intake failed: {e}')
    rec=[]
    for f in files or []:rec.append({'name':f.name,'data':f.getvalue(),'source':'Individual upload'})
    if zup:
        try:
            with zipfile.ZipFile(io.BytesIO(zup.getvalue())) as z:
                for i in z.infolist():
                    if not i.is_dir() and Path(i.filename).suffix.lower() in SUPPORTED:rec.append({'name':Path(i.filename).name,'data':z.read(i),'source':'ZIP'})
        except Exception as e:errs.append(f'ZIP intake failed: {e}')
    for u in parse_links(links):
        try:
            x=download(u);rec.append({'name':'linked.pdf' if x.startswith(b'%PDF') else 'linked.docx','data':x,'source':'Public link'})
        except Exception as e:errs.append(f'{u}: {e}')
    if not rec:errs.append('No CVs were successfully received.')
    for e in errs:st.warning(e)
    if jd and rec:
        seen=set();fps=set();unique=[];dups=[]
        for r in rec:
            if h(r['data']) in seen:dups.append((r,'Exact duplicate file'));continue
            seen.add(h(r['data']))
            try:t=extract(r['data'],r['name']);fp=h(re.sub(r'\W+','',t.lower()[:20000]).encode()) if t else h(r['data'])
            except:t='';fp=h(r['data'])
            if fp in fps:dups.append((r,'Duplicate document content'));continue
            fps.add(fp);r['text']=t;unique.append(r)
        results=[];failed=[]
        for r in unique:
            try:
                if len(r['text'])<40:raise ValueError('Too little readable text extracted')
                results.append(score(r['text'],r['name'],jd))
            except Exception as e:failed.append({'File':r['name'],'Reason':str(e)})
        if results:
            rows=[]
            for r in results:
                yes=[x['Requirement'] for x in r['Evidence Matrix'] if x['Match']=='Yes'][:3];gap=[x['Requirement'] for x in r['Evidence Matrix'] if x['Match']!='Yes'][:3]
                rows.append({'Name':r['Name'],'Fit %':r['Fit %'],'Ranking Group':r['Ranking Group'],'2-Line Verdict':r['2-Line Verdict'],'Why Not 100%':r['Why Not 100%'],'Top 3 Matching Skills':'; '.join(yes) or 'None explicitly evidenced','Top Gaps':'; '.join(gap) or 'None material','Years Exp':r['Years Exp'],'Location':r['Location']})
            df=pd.DataFrame(rows).sort_values(['Fit %','Name'],ascending=[False,True]).reset_index(drop=True);meta={'Prepared exclusively for':cn or 'Client','Client Address':ca or 'Not provided','Client Contact Officer':co or 'Not provided','Client Contact Email':ce or 'Not provided','Recruitment Project':role or 'Not provided','Screening date':str(sd)};st.session_state.result={'df':df,'meta':meta,'dups':dups,'failed':failed}
if st.session_state.result:
    R=st.session_state.result;df=R['df'];meta=R['meta'];st.success(f'Screening complete: {len(df)} CVs scored.')
    if R['dups']:st.warning(f'{len(R["dups"])} duplicate CV(s) excluded.')
    if R['failed']:st.warning(f'{len(R["failed"])} CV(s) could not be read.')
    st.dataframe(df,use_container_width=True);a,b,c=st.tabs(['Ranking Distribution','Score Distribution','Top Candidates'])
    with a:st.image(charts(df,'pie'))
    with b:st.image(charts(df,'hist'))
    with c:st.image(charts(df,'top'))
    if R['dups']:
        with st.expander('Duplicate Entry Detection Report'):st.dataframe(pd.DataFrame([{'File':x['name'],'Source':x['source'],'Reason':why} for x,why in R['dups']]))
    if R['failed']:
        with st.expander('Unreadable / Failed Files'):st.dataframe(pd.DataFrame(R['failed']))
    x1,x2,x3,x4=st.columns(4)
    with x1:st.download_button('Download Workbook',excel(df,meta),'screening_results.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with x2:st.download_button('Download CSV',df.to_csv(index=False).encode(),'screening_results.csv','text/csv')
    with x3:st.download_button('Download Professional DOCX',docx(df,meta),'candidate_screening_report.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    with x4:st.download_button('Download Professional PDF',pdf(df,meta),'candidate_screening_report.pdf','application/pdf')
