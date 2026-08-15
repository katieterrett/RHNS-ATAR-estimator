import camelot, re, pandas as pd, numpy as np
from pathlib import Path

BASE=Path('/mnt/data')
OUT=Path('/mnt/data/atar_estimator_5yr')
page_ranges={2021:'46-50',2022:'46-50',2023:'47-51',2024:'47-52',2025:'47-51'}
normalize={
    'English EALD':'English EAL/D',
}
rows=[]
for year,pages in page_ranges.items():
    tables=camelot.read_pdf(str(BASE/f'uac{year}.pdf'), pages=pages, flavor='stream')
    pending=None
    for tb in tables:
        for _,row in tb.df.iterrows():
            cells=[str(x).strip() for x in row.tolist()]
            idx=next((i for i,c in enumerate(cells) if c.lower()=='hsc'),None)
            if idx is not None:
                course=' '.join([c for c in cells[:idx-1] if c]).strip()
                number=cells[idx-1].replace(',','') if idx>=1 else ''
                vals=[]
                for c in cells[idx+1:]:
                    try: vals.append(float(c.replace(',','')))
                    except Exception: pass
                pending={'year':year,'course':normalize.get(course,course),'number':number,'hsc':vals[:8]} if len(vals)>=8 and course else None
                continue
            sidx=next((i for i,c in enumerate(cells) if c.lower() in ('scaled','sca')),None)
            if sidx is not None and pending:
                vals=[]
                for c in cells[sidx+1:]:
                    try: vals.append(float(c.replace(',','')))
                    except Exception: pass
                if len(vals)>=8:
                    pending['scaled']=vals[:8]
                    rows.append(pending)
                pending=None
            elif pending:
                cont=' '.join([c for c in cells[:3] if c and not re.fullmatch(r'[\d,]+',c)]).strip()
                blocked=['table a3','appendix','course number','type of','mark mean','descriptive statistics']
                if cont and not any(x in cont.lower() for x in blocked):
                    pending['course']=(pending['course']+' '+cont).strip()

out=[]
labels=['mean','sd','max','p99','p90','p75','p50','p25']
for r in rows:
    rec={'year':r['year'],'course':r['course'],'number':int(r['number']) if str(r['number']).isdigit() else np.nan}
    for lab,val in zip(labels,r['hsc']): rec[f'hsc_{lab}']=val
    for lab,val in zip(labels,r['scaled']): rec[f'scaled_{lab}']=val
    out.append(rec)

df=pd.DataFrame(out).drop_duplicates(['year','course'])
# Clean current/legacy name continuation anomalies if any.
df['course']=df['course'].str.replace(r'\s+',' ',regex=True).str.strip()
df.to_csv(OUT/'uac_scaling_points_2021_2025.csv',index=False)

# Aggregate thresholds from UAC 2025 Table A9 (contains 2021-2025).
atar=[99.95,99.50,99.00,98.00,95.00,90.00,85.00,80.00,75.00,70.00,65.00,60.00,55.00,50.00]
vals={
2021:[478.8,458.7,448.0,433.7,404.8,369.8,340.2,313.5,288.0,263.8,239.8,217.2,195.4,172.8],
2022:[478.1,459.9,449.1,434.3,404.6,368.9,338.9,310.9,285.2,259.5,234.4,210.4,186.4,162.8],
2023:[479.8,459.3,449.5,435.0,405.5,370.0,340.2,313.0,287.4,261.8,236.8,212.5,188.2,164.3],
2024:[477.4,455.9,445.6,431.6,403.5,369.2,340.2,312.6,286.2,260.6,235.4,210.1,185.3,160.6],
2025:[479.4,459.4,447.8,433.2,404.6,370.7,341.6,313.8,287.5,262.0,237.2,212.9,190.0,165.7],
}
agg=[]
for y,xs in vals.items():
    for a,x in zip(atar,xs): agg.append({'year':y,'atar':a,'aggregate':x})
pd.DataFrame(agg).to_csv(OUT/'uac_atar_aggregate_2021_2025.csv',index=False)

print(df.groupby('year').size().to_string())
print('Unique courses:',df.course.nunique())
