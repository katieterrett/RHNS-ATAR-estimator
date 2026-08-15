from __future__ import annotations
from pathlib import Path
import math
import numpy as np
import pandas as pd

HERE=Path(__file__).resolve().parent
SCALING=pd.read_csv(HERE/'uac_scaling_points_2021_2025.csv')
AGG=pd.read_csv(HERE/'uac_atar_aggregate_2021_2025.csv')
YEARS=sorted(SCALING.year.unique().tolist())


def course_units(course: str) -> int:
    if 'Extension' in course or course == 'Studies of Religion I':
        return 1
    return 2


def is_english(course: str) -> bool:
    return course.startswith('English ')


def available_courses() -> list[str]:
    return sorted(SCALING.course.unique().tolist())


def _curve(course: str, year: int):
    r=SCALING[(SCALING.course==course)&(SCALING.year==year)]
    if r.empty:
        return None
    r=r.iloc[0]
    names=['p25','p50','p75','p90','p99','max']
    pairs=[]
    for n in names:
        x=float(r[f'hsc_{n}']); y=float(r[f'scaled_{n}'])
        if np.isfinite(x) and np.isfinite(y): pairs.append((x,y))
    # Sort, then collapse duplicate HSC x values using the largest scaled value.
    byx={}
    for x,y in pairs:
        byx[x]=max(y,byx.get(x,-1e9))
    pts=sorted(byx.items())
    return pts if len(pts)>=2 else None


def _linear_interp_extrap(x: float, pts: list[tuple[float,float]], lo=0.0, hi=50.0) -> float:
    pts=sorted(pts)
    xs=np.array([p[0] for p in pts],dtype=float)
    ys=np.array([p[1] for p in pts],dtype=float)
    if x <= xs[0]:
        x0,x1=xs[0],xs[1]; y0,y1=ys[0],ys[1]
        y=y0+(x-x0)*(y1-y0)/(x1-x0)
    elif x >= xs[-1]:
        x0,x1=xs[-2],xs[-1]; y0,y1=ys[-2],ys[-1]
        y=y1+(x-x1)*(y1-y0)/(x1-x0)
    else:
        y=float(np.interp(x,xs,ys))
    return float(np.clip(y,lo,hi))


def scaled_per_unit(course: str, predicted_hsc_100: float, year: int) -> float | None:
    """Estimate UAC scaled mark per unit (0-50) from a predicted HSC-equivalent mark on a 0-100 scale."""
    pts=_curve(course,year)
    if not pts:
        return None
    return _linear_interp_extrap(float(predicted_hsc_100)/2.0,pts)


def mean_scaled_per_unit(course: str, predicted_hsc_100: float) -> tuple[float|None,dict[int,float]]:
    vals={}
    for y in YEARS:
        v=scaled_per_unit(course,predicted_hsc_100,y)
        if v is not None: vals[y]=v
    if not vals: return None,{}
    return float(np.mean(list(vals.values()))),vals


def _atar_curve_for_year(year: int) -> list[tuple[float,float]]:
    d=AGG[AGG.year==year].sort_values('aggregate')
    return list(zip(d['aggregate'].astype(float),d['atar'].astype(float)))


def mean_atar_curve() -> list[tuple[float,float]]:
    p=AGG.pivot(index='atar',columns='year',values='aggregate')
    mean=p.mean(axis=1).reset_index(name='aggregate').sort_values('aggregate')
    return list(zip(mean['aggregate'].astype(float),mean['atar'].astype(float)))


def atar_from_aggregate(aggregate: float, year: int|None=None) -> float | None:
    pts=_atar_curve_for_year(year) if year is not None else mean_atar_curve()
    xs=np.array([p[0] for p in pts]); ys=np.array([p[1] for p in pts])
    if aggregate < xs.min():
        return None  # UAC Table A9 published here starts at ATAR 50.
    if aggregate >= xs.max():
        return 99.95
    val=float(np.interp(aggregate,xs,ys))
    return round(np.clip(round(val/0.05)*0.05,50,99.95)+1e-10,2)


def select_best_10(per_unit_by_course: dict[str,float]) -> tuple[float,dict[str,int],list[dict]]:
    units=[]
    for course,scaled in per_unit_by_course.items():
        for i in range(course_units(course)):
            units.append({'course':course,'scaled':float(scaled),'english':is_english(course),'unit_index':i+1})
    english=sorted([u for u in units if u['english']],key=lambda u:u['scaled'],reverse=True)
    if len(english)<2 or len(units)<10:
        raise ValueError('At least 10 units including at least 2 units of English are required.')
    chosen=english[:2]
    chosen_ids={(u['course'],u['unit_index']) for u in chosen}
    remaining=[u for u in units if (u['course'],u['unit_index']) not in chosen_ids]
    chosen += sorted(remaining,key=lambda u:u['scaled'],reverse=True)[:8]
    counts={}
    for u in chosen: counts[u['course']]=counts.get(u['course'],0)+1
    return sum(u['scaled'] for u in chosen),counts,chosen


def estimate(rows: list[dict], sensitivity: float=2.0) -> dict:
    """rows: [{'course': str, 'mark': 0..100}, ...]"""
    clean=[]; seen=set()
    for r in rows:
        c=str(r['course']).strip(); m=float(r['mark'])
        if not c or c in seen: continue
        if c not in available_courses(): raise ValueError(f'Unknown course: {c}')
        if not 0<=m<=100: raise ValueError('Predicted HSC marks must be between 0 and 100.')
        seen.add(c); clean.append({'course':c,'mark':m})
    if sum(course_units(r['course']) for r in clean)<10:
        raise ValueError('Enter at least 10 units.')
    if sum(course_units(r['course']) for r in clean if is_english(r['course']))<2:
        raise ValueError('Enter at least 2 units of English.')

    subject=[]; central_map={}
    yearly_maps={y:{} for y in YEARS}
    for r in clean:
        mean,vals=mean_scaled_per_unit(r['course'],r['mark'])
        if mean is None: continue
        central_map[r['course']]=mean
        for y,v in vals.items(): yearly_maps[y][r['course']]=v
        contrib=atar_from_aggregate(mean*10)
        subject.append({
            'course':r['course'],'predicted_hsc':r['mark'],'units':course_units(r['course']),
            'mean_scaled_per_unit':mean,'scaled_min':min(vals.values()),'scaled_max':max(vals.values()),
            'years_available':len(vals),'atar_contribution':contrib,
            **{f'scaled_{y}': vals.get(y,np.nan) for y in YEARS}
        })
    if len(central_map)!=len(clean):
        missing=[r['course'] for r in clean if r['course'] not in central_map]
        raise ValueError('No usable scaling data for: '+', '.join(missing))

    central_agg,counts,chosen=select_best_10(central_map)
    central_atar=atar_from_aggregate(central_agg)

    yearly=[]
    course_set=set(central_map)
    for y in YEARS:
        if set(yearly_maps[y])==course_set:
            agg,_,_=select_best_10(yearly_maps[y])
            yearly.append({'year':y,'aggregate':agg,'atar':atar_from_aggregate(agg,y)})

    def scenario(delta):
        m={}
        for r in clean:
            mean,_=mean_scaled_per_unit(r['course'],min(100,max(0,r['mark']+delta)))
            m[r['course']]=mean
        agg,_,_=select_best_10(m)
        return {'aggregate':agg,'atar':atar_from_aggregate(agg)}

    low=scenario(-abs(sensitivity)); high=scenario(abs(sensitivity))
    return {
        'central_aggregate':central_agg,'central_atar':central_atar,'counts':counts,'chosen_units':chosen,
        'subjects':subject,'yearly':yearly,'low':low,'high':high,'sensitivity':abs(sensitivity),
        'years':YEARS,
    }
