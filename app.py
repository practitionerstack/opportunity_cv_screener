import io, re, zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title='Opportunity Hub CV Screener v2', page_icon='🎯', layout='wide')

SKILLS = {
'Meta Ads':['meta ads','facebook ads','instagram ads','meta advertising'],
'Google Ads':['google ads','google adwords','adwords'],
'GA4':['ga4','google analytics 4','google analytics'],
'Email Marketing':['email marketing','email campaigns','email campaign','newsletter','newsletters','mailchimp'],
'Canva':['canva'],
'Copywriting':['copywriting','copywriter','ad copy','social copy','marketing copy'],
'HubSpot':['hubspot'],
'A/B Testing':['a/b testing','ab testing','split testing'],
'Google Tag Manager':['google tag manager','gtm'],
}
FINTECH = ['kuda','opay','o-pay','carbon','flutterwave','paystack','moniepoint','paga','gtbank','guaranty trust bank','access bank','zenith bank','first bank','uba','stanbic ibtc','wema bank','palmpay']


def norm(x): return re.sub(r'\s+',' ',str(x).lower()).strip()

def read_text_file(f):
    name=f.name.lower(); data=f.getvalue()
    if name.endswith('.pdf'):
        return '\n'.join((p.extract_text() or '') for p in PdfReader(io.BytesIO(data)).pages)
    if name.endswith('.docx'):
        d=Document(io.BytesIO(data)); return '\n'.join(p.text for p in d.paragraphs)
    if name.endswith('.txt'): return data.decode('utf-8',errors='ignore')
    return ''

def read_structured_cvs(f):
    data=f.getvalue(); name=f.name.lower()
    if name.endswith('.csv'): df=pd.read_csv(io.BytesIO(data))
    else: df=pd.read_excel(io.BytesIO(data))
    cols={c.lower():c for c in df.columns}
    textcol=next((cols[x] for x in cols if x in ['cv','cv text','resume','resume text','text','content']),None)
    namecol=next((cols[x] for x in cols if x in ['name','candidate','candidate name','full name']),None)
    if not textcol: raise ValueError('Spreadsheet needs a CV/Resume/Text column containing candidate CV text.')
    out=[]
    for i,r in df.fillna('').iterrows():
        txt=str(r[textcol]); nm=str(r[namecol]) if namecol else f'{f.name} row {i+1}'
        if len(txt.strip())>=20: out.append((nm,txt,f'{f.name} row {i+1}'))
    return out

def extract_name(text, fallback):
    for line in [x.strip() for x in text.splitlines() if x.strip()][:12]:
        if re.fullmatch(r'[A-Za-z .\-]{5,60}',line) and not any(k in line.lower() for k in ['curriculum','resume','cv','experience','marketing','executive','specialist','lead','skills','education','profile']): return line.title()
    return re.sub(r'[_\-]+',' ',fallback.rsplit('.',1)[0]).title()

def years_exp(text):
    spans=[]; now=datetime.now().year
    for a,b in re.findall(r'\b(20\d{2})\s*[-–]\s*(20\d{2}|present|current)\b',text,re.I):
        a=int(a); b=now if b.lower() in ['present','current'] else int(b)
        if b>=a: spans.append((a,b))
    if not spans:return 0
    yrs=set()
    for a,b in spans: yrs.update(range(a,b+1))
    return max(0,len(yrs)-1)

def explicit_negative(t, alias):
    a=re.escape(alias)
    pats=[rf'\bno\s+{a}\b',rf'\bno\s+experience\s+(?:in|with)\s+{a}\b',rf'\b{a}\s*:\s*no\b']
    return any(re.search(p,t,re.I) for p in pats)

def evidence(text, skill):
    t=norm(text); aliases=SKILLS.get(skill,[skill.lower()])
    for a in aliases:
        if explicit_negative(t,a): return 'missing'
    hits=[]
    for a in aliases:
        hits += [m.start() for m in re.finditer(re.escape(a),t,re.I)]
    if not hits:return 'unknown'
    for p in hits:
        w=t[max(0,p-100):p+180]
        if any(x in w for x in ['boosted post','basic ','assisted','support','exposure to','intern','helped with']): return 'partial'
    return 'confirmed'

def fintech_evidence(text):
    t=norm(text)
    return next((x for x in FINTECH if re.search(r'(?<![a-z])'+re.escape(x)+r'(?![a-z])',t)),None)

def location(text):
    t=norm(text)
    if 'lagos' in t:return 'Lagos / Strong'
    if 'remote' in t:return 'Remote / Review'
    for c in ['abuja','ibadan','port harcourt','enugu','kano','benin','kaduna']:
        if c in t:return f'{c.title()} / Review'
    return 'Unknown / Review'

def parse_jd(text,title):
    j=norm(text); req=[]
    for s,aliases in SKILLS.items():
        if any(a in j for a in aliases):
            # preferred/nice gets lower weight
            pos=min([j.find(a) for a in aliases if a in j])
            ctx=j[max(0,pos-80):pos+100]
            req.append((s,5 if any(x in ctx for x in ['nice to have','nice-to-have','preferred','plus']) else 12))
    if not req and any(x in j for x in ['digital marketing','performance marketing','marketing executive']):
        req=[('Meta Ads',12),('Google Ads',12),('GA4',10),('Email Marketing',10),('Canva',5),('Copywriting',8)]
    return {'title':title,'text':text,'requirements':req,'min_years': int(re.search(r'(\d+)\s*\+?\s*years',j).group(1)) if re.search(r'(\d+)\s*\+?\s*years',j) else 0,'industry':('fintech/banking' if any(x in j for x in ['fintech','banking']) else ''),'lagos_hybrid':('hybrid' in j and 'lagos' in j)}

def score(text,jd):
    deductions=[]; gaps=[]; matched=[]; audit=[]; score=100
    for skill,w in jd['requirements']:
        ev=evidence(text,skill)
        if ev=='confirmed': matched.append(skill); audit.append(f'{skill}: full evidence')
        elif ev=='partial':
            d=max(1,w//2); score-=d; deductions.append(f'Limited {skill}: -{d}pts'); gaps.append(f'Limited {skill}'); matched.append(f'{skill} (basic)'); audit.append(f'{skill}: partial evidence')
        else:
            score-=w; deductions.append(f'No confirmed {skill}: -{w}pts'); gaps.append(f'No confirmed {skill}'); audit.append(f'{skill}: no evidence')
    yrs=years_exp(text)
    if jd['min_years'] and yrs<jd['min_years']:
        d=15; score-=d; deductions.append(f'Below {jd["min_years"]}+ years: -{d}pts'); gaps.append('Below minimum experience')
    elif yrs: audit.append(f'Experience: {yrs} years')
    ft=fintech_evidence(text)
    if jd['industry']:
        if ft: audit.append(f'Industry: confirmed ({ft})')
        else: score-=8; deductions.append('No confirmed fintech/banking background: -8pts'); gaps.append('No confirmed industry background')
    loc=location(text)
    if jd['lagos_hybrid'] and not loc.startswith('Lagos'):
        score-=5; deductions.append('Not Lagos-based for hybrid role: -5pts'); gaps.append('Location mismatch')
    score=max(0,score)
    if score>=90: group='Excellent'; verdict='Excellent match; shortlist immediately.'
    elif score>=70: group='Good'; verdict='Strong candidate; shortlist with minor gaps.'
    elif score>=50: group='Moderate'; verdict='Possible fit; gaps require careful review.'
    elif score>=30: group='Maybe'; verdict='Weak fit; interview only with strong compensating evidence.'
    else: group='Do Not Hire'; verdict='Poor match; do not prioritize for shortlist.'
    green=[]
    if ft: green.append(f'Confirmed fintech/banking background ({ft})')
    if yrs: green.append(f'{yrs}+ years experience')
    green += [x for x in matched if '(basic)' not in x]
    return score,group,verdict,'; '.join(deductions[:6]) or 'Meets stated requirements','; '.join(gaps[:2]) or 'None material','; '.join(green[:3]) or 'Relevant transferable evidence',yrs,', '.join(matched) or 'None confirmed',loc,'; '.join(audit)

def make_excel(df):
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        df.to_excel(w,index=False,sheet_name='Summary')
        for g in ['Excellent','Good','Moderate','Maybe','Do Not Hire','Review Required']:
            x=df[df['Ranking Group']==g]
            x.to_excel(w,index=False,sheet_name=g[:31])
    return out.getvalue()

def make_docx(df,jobs):
    d=Document(); d.add_heading('Opportunity Hub Candidate Screening Summary',0)
    d.add_paragraph(f'Screening date: {datetime.now().strftime("%d %B %Y")}')
    d.add_paragraph(f'Job descriptions processed: {len(jobs)}')
    d.add_paragraph(f'Candidate records screened: {len(df)}')
    d.add_heading('Executive Summary',1)
    for g,n in df['Ranking Group'].value_counts().reindex(['Excellent','Good','Moderate','Maybe','Do Not Hire'],fill_value=0).items(): d.add_paragraph(f'{g}: {n}')
    d.add_heading('Top Recommendations',1)
    for _,r in df.sort_values('Fit %',ascending=False).head(10).iterrows(): d.add_paragraph(f"{r['Name']} — {r['Fit %']}% — {r['Best Job Match']}")
    d.add_heading('Important Note',1); d.add_paragraph('This report is decision support. Human review is required before employment decisions.')
    b=io.BytesIO(); d.save(b); return b.getvalue()

st.title('🎯 Opportunity Hub — CV Screening Service v2')
st.caption('Evidence-based candidate screening with single or multi-job matching, grouped exports and executive summary reports.')

with st.sidebar:
    mode=st.radio('Job intake mode',['Paste / Upload JD','Direct role builder'])
    st.info('For very large volumes, use structured CSV/XLSX exports or batch files. Million-scale processing requires the future background-worker infrastructure.')

jobs=[]
if mode=='Paste / Upload JD':
    pasted=st.text_area('Paste Job Description (optional)',height=180)
    jd_files=st.file_uploader('Upload one or more Job Descriptions',type=['pdf','docx','txt'],accept_multiple_files=True)
    if pasted.strip(): jobs.append(parse_jd(pasted,'Pasted Job Description'))
    for f in jd_files or []:
        try:
            txt=read_text_file(f)
            if len(txt.strip())>=20: jobs.append(parse_jd(txt,f.name))
        except Exception as e: st.error(f'Could not read {f.name}: {e}')
else:
    title=st.text_input('Job role title',placeholder='e.g. Digital Marketing Executive')
    industry=st.text_input('Industry (optional)',placeholder='e.g. Fintech')
    minyrs=st.number_input('Minimum years experience',0,30,2)
    selected=st.multiselect('Required / preferred skills',list(SKILLS.keys()))
    location_req=st.selectbox('Work/location requirement',['Any location','Lagos hybrid','Remote','Lagos'])
    if title.strip():
        txt=f'{title}. {industry}. {minyrs}+ years. Skills: '+', '.join(selected)+'. '+location_req
        jobs=[parse_jd(txt,title)]

cv_files=st.file_uploader('Upload CVs or CV data files',type=['pdf','docx','txt','csv','xlsx','xls'],accept_multiple_files=True)
run=st.button('Screen Candidates',type='primary',disabled=not(jobs and cv_files))

if run:
    candidates=[]
    for f in cv_files:
        try:
            if f.name.lower().endswith(('.csv','.xlsx','.xls')): candidates.extend(read_structured_cvs(f))
            else:
                txt=read_text_file(f)
                if len(txt.strip())<20: raise ValueError('Could not extract enough text')
                candidates.append((extract_name(txt,f.name),txt,f.name))
        except Exception as e:
            candidates.append((f.name,'',f.name+f' [EXTRACTION FAILED: {str(e)[:100]}]'))
    rows=[]; bar=st.progress(0); status=st.empty()
    total=len(candidates)
    for i,(name,text,source) in enumerate(candidates):
        if not text:
            rows.append({'Name':name,'Fit %':0,'2-Line Verdict':'Could not reliably process this CV.','Why Not 100%':'Extraction failed','Red Flag':'Extraction failed','Green Flag':'Manual review required','Years Exp':0,'Skills Match':'','Visa/Location':'Unknown','Best Job Match':'','Ranking Group':'Review Required','Resume Link':source,'Audit Trail':'Extraction failed'})
        else:
            options=[]
            for j in jobs:
                s=score(text,j); options.append((s[0],j['title'],s))
            best=max(options,key=lambda x:x[0]); s=best[2]
            rows.append({'Name':name,'Fit %':s[0],'2-Line Verdict':s[2],'Why Not 100%':s[3],'Red Flag':s[4],'Green Flag':s[5],'Years Exp':s[6],'Skills Match':s[7],'Visa/Location':s[8],'Best Job Match':best[1],'Ranking Group':s[1],'Resume Link':source,'Audit Trail':s[9]})
        bar.progress((i+1)/max(1,total)); status.caption(f'Processed {i+1:,} of {total:,} candidate records')
    st.session_state.results=pd.DataFrame(rows).sort_values(['Fit %','Name'],ascending=[False,True]).reset_index(drop=True)
    st.session_state.jobs=jobs

if 'results' in st.session_state:
    df=st.session_state.results; jobs=st.session_state.jobs
    st.subheader(f'Ranked Results — {len(df):,} Candidates')
    st.dataframe(df[['Name','Fit %','2-Line Verdict','Why Not 100%','Red Flag','Green Flag','Years Exp','Skills Match','Visa/Location','Resume Link']],use_container_width=True,hide_index=True)
    st.download_button('⬇ Download grouped Excel workbook',make_excel(df),'opportunity_hub_screening.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.download_button('⬇ Download CSV',df.to_csv(index=False).encode(),'opportunity_hub_screening.csv','text/csv')
    st.download_button('⬇ Download executive summary DOCX',make_docx(df,jobs),'opportunity_hub_screening_summary.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document')
