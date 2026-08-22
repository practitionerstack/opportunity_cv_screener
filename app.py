import io, re, zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title='Opportunity Hub CV Screener', page_icon='🎯', layout='wide')

SKILL_MAP = {
    'Meta Ads': ['meta ads','facebook ads','instagram ads','meta advertising'],
    'Google Ads': ['google ads','adwords'],
    'GA4': ['ga4','google analytics 4','google analytics'],
    'Email Marketing': ['email marketing','email campaign','newsletter','mailchimp','hubspot'],
    'Canva': ['canva'],
    'Copywriting': ['copywriting','copywriter','ad copy','social copy'],
    'HubSpot': ['hubspot'],
    'A/B Testing': ['a/b testing','ab testing','split testing'],
    'Google Tag Manager': ['google tag manager','gtm'],
}

REQUIRED = ['Meta Ads','Google Ads','GA4','Email Marketing','Canva','Copywriting']
NICE = ['HubSpot']

def read_file(uploaded):
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    if name.endswith('.pdf'):
        return '\n'.join((p.extract_text() or '') for p in PdfReader(io.BytesIO(data)).pages)
    if name.endswith('.docx'):
        doc = Document(io.BytesIO(data)); return '\n'.join(p.text for p in doc.paragraphs)
    if name.endswith('.txt'):
        return data.decode('utf-8', errors='ignore')
    if name.endswith('.csv'):
        return pd.read_csv(io.BytesIO(data)).astype(str).to_csv(index=False)
    if name.endswith(('.xlsx','.xls')):
        xls = pd.ExcelFile(io.BytesIO(data)); return '\n'.join(pd.read_excel(xls, s).astype(str).to_csv(index=False) for s in xls.sheet_names)
    return ''

def norm(s): return re.sub(r'\s+',' ',s.lower())

def first_match(patterns, text):
    t=norm(text)
    return any(p in t for p in patterns)

def extract_name(text, filename):
    lines=[x.strip() for x in text.splitlines() if x.strip()]
    for line in lines[:8]:
        if re.fullmatch(r'[A-Za-z .\-]{5,60}', line) and not any(k in line.lower() for k in ['curriculum','resume','cv','experience','marketing','executive','specialist','lead']):
            return line.title()
    return re.sub(r'[_\-]+',' ',filename.rsplit('.',1)[0]).title()

def years_experience(text):
    years=set()
    for a,b in re.findall(r'\b(20\d{2})\s*[-–]\s*(20\d{2}|present|current)\b', text, re.I):
        start=int(a); end=datetime.now().year if b.lower() in ('present','current') else int(b)
        if end>=start: years.update(range(start,end+1))
    return max(0,len(years)-1)

def has_fintech(text):
    return any(x in norm(text) for x in ['bank','fintech','kuda','opay','carbon','flutterwave','gtbank','access bank','moniepoint','paystack'])

def location(text):
    t=norm(text)
    if 'lagos' in t: return 'Lagos / Strong'
    if 'remote' in t: return 'Remote / Review'
    for city in ['abuja','ibadan','port harcourt','enugu','kano']:
        if city in t: return f'{city.title()} / Review'
    return 'Unknown / Review'

def score_cv(text, jd):
    t=norm(text); j=norm(jd)
    name_skills=[]
    for skill, patterns in SKILL_MAP.items():
        if first_match(patterns,t): name_skills.append(skill)
    yrs=years_experience(text)
    score=100; deductions=[]; gaps=[]
    # Exact role requirements, strict but capped.
    req_checks=[
        ('Meta Ads','No Meta Ads experience',12),('Google Ads','No Google Ads experience',12),
        ('GA4','No GA4/Google Analytics',10),('Email Marketing','No email marketing experience',10),
        ('Canva','No Canva experience',5),('Copywriting','No copywriting evidence',8)]
    for skill,msg,pts in req_checks:
        if skill not in name_skills: score-=pts; deductions.append(f'{msg}: -{pts}pts'); gaps.append(msg)
    if yrs < 2:
        score-=15; deductions.append(f'Less than 2 years relevant experience: -15pts'); gaps.append('Below minimum experience')
    if not has_fintech(text):
        score-=8; deductions.append('No fintech/banking experience: -8pts'); gaps.append('No fintech/banking background')
    if 'lagos' not in t:
        score-=5; deductions.append('Not Lagos-based for hybrid role: -5pts'); gaps.append('Location mismatch')
    # Cap penalty explanation to most material items; never below 0.
    score=max(0,score)
    green=[]
    if has_fintech(text): green.append('Direct fintech/banking experience')
    if yrs>=2: green.append(f'{yrs}+ years experience')
    for x in ['Meta Ads','Google Ads','GA4','Email Marketing','Copywriting']:
        if x in name_skills and len(green)<3: green.append(x)
    red='None material' if score>=80 else (gaps[0] if gaps else 'Review evidence')
    if score>=90: verdict='Excellent match; shortlist immediately.'
    elif score>=70: verdict='Strong candidate; shortlist with minor gaps.'
    elif score>=50: verdict='Possible fit; gaps require careful review.'
    else: verdict='Poor match; do not prioritize for shortlist.'
    return {
        'Fit %':score,'2-Line Verdict':verdict,
        'Why Not 100%':'; '.join(deductions[:3]) or 'Meets stated requirements',
        'Red Flag':red,'Green Flag':'; '.join(green[:3]) or 'Some transferable experience',
        'Years Exp':yrs,'Skills Match':', '.join(name_skills) or 'None detected',
        'Visa/Location':location(text),
        '_gaps':'; '.join(gaps[:3])
    }

st.title('🎯 Opportunity Hub — AI CV Screening Service')
st.caption('Upload a Job Description and multiple CVs. The engine extracts, scores, ranks, and exports candidates.')

with st.sidebar:
    st.header('Scoring mode')
    st.info('Strict Nigerian HR screening MVP. Scores are rule-based for consistency; evidence is shown for review.')

jd = st.text_area('Job Description', height=220, placeholder='Paste the complete Job Description here...')
files = st.file_uploader('Upload CVs', type=['pdf','docx','txt','csv','xlsx','xls'], accept_multiple_files=True)

if st.button('Screen Candidates', type='primary', disabled=not (jd and files)):
    rows=[]
    progress=st.progress(0)
    for i,f in enumerate(files):
        try:
            text=read_file(f)
            if len(text.strip())<20: raise ValueError('Could not extract enough text')
            result=score_cv(text,jd)
            rows.append({'Name':extract_name(text,f.name), **result, 'Resume Link':f.name})
        except Exception as e:
            rows.append({'Name':f.name,'Fit %':0,'2-Line Verdict':'Could not reliably process this CV.','Why Not 100%':str(e)[:120],'Red Flag':'Extraction failed','Green Flag':'Manual review required','Years Exp':0,'Skills Match':'','Visa/Location':'Unknown','Resume Link':f.name})
        progress.progress((i+1)/len(files))
    df=pd.DataFrame(rows).sort_values('Fit %',ascending=False).reset_index(drop=True)
    st.session_state['results']=df

if 'results' in st.session_state:
    df=st.session_state['results']
    st.subheader(f'Ranked Results — {len(df)} Candidates')
    display=df[['Name','Fit %','2-Line Verdict','Why Not 100%','Red Flag','Green Flag','Years Exp','Skills Match','Visa/Location','Resume Link']]
    st.dataframe(display, use_container_width=True, hide_index=True)
    out=io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        display.to_excel(writer,index=False,sheet_name='Screening Results')
    st.download_button('⬇ Download Excel Report',out.getvalue(),'opportunity_hub_screening.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.download_button('⬇ Download CSV',display.to_csv(index=False).encode(),'opportunity_hub_screening.csv','text/csv')
