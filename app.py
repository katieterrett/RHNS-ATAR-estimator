from __future__ import annotations

from datetime import date
from pathlib import Path
import re
import pandas as pd
import streamlit as st

from model import estimate, available_courses
from export_excel import build_excel_export

SCHOOL_NAME = 'Reddam House North Shore'
NAVY = '#1B293B'
GOLD = '#A68E3A'
PALE = '#F6F3EB'
LOGO_PATH = Path(__file__).resolve().parent / 'assets' / 'reddam_house_north_shore_logo.png'

st.set_page_config(
    page_title=f'{SCHOOL_NAME} | ATAR Estimator',
    page_icon='🎓',
    layout='wide',
)

# Reddam House North Shore branded presentation.
st.markdown(f'''
<style>
    .stApp {{ background: #FBFBFA; }}
    .block-container {{ padding-top: 1.0rem; padding-bottom: 3rem; max-width: 1450px; }}
    .rh-brandbar {{
        background: #FFFFFF;
        border: 1px solid #E1E5E9;
        border-bottom: 4px solid {GOLD};
        padding: .65rem 1rem;
        border-radius: 10px 10px 0 0;
        margin-bottom: 0;
    }}
    .rh-header {{
        background: linear-gradient(120deg, {NAVY} 0%, #1C314B 100%);
        padding: 1rem 1.7rem 1.05rem 1.7rem;
        border-radius: 0 0 10px 10px;
        margin-bottom: 1rem;
        box-shadow: 0 4px 14px rgba(20,36,58,.12);
    }}
    .rh-title {{ color: white; font-size: 2.0rem; font-weight: 650; margin: .1rem 0 0 0; }}
    .rh-subtitle {{ color: #DCE3EB; font-size: .95rem; margin-top: .35rem; }}
    .rh-kicker {{ color: {GOLD}; font-weight: 700; text-transform: uppercase; letter-spacing: .09em; font-size: .75rem; }}
    div[data-testid="stMetric"] {{
        background: white;
        border: 1px solid #E1E5E9;
        border-top: 3px solid {GOLD};
        padding: 0.7rem 0.8rem;
        border-radius: 8px;
    }}
    div[data-testid="stDataFrame"] {{ border: 1px solid #E1E5E9; border-radius: 8px; overflow: hidden; }}
    .rh-note {{ background: {PALE}; border-left: 4px solid {GOLD}; padding: .8rem 1rem; border-radius: 4px; color: #334155; }}
    h2, h3 {{ color: {NAVY}; }}
</style>
''', unsafe_allow_html=True)

# Official supplied school logo. It is kept on the page's white background so the navy wordmark remains legible.
if LOGO_PATH.exists():
    st.image(str(LOGO_PATH), width=500)
else:
    st.markdown('**REDDAM HOUSE NORTH SHORE**')

st.markdown('''
<div class="rh-header">
  <div class="rh-kicker">Academic Analytics</div>
  <div class="rh-title">NSW HSC → ATAR Estimator</div>
  <div class="rh-subtitle">Five-year model using published UAC scaling and aggregate data, 2021–2025</div>
</div>
''', unsafe_allow_html=True)

with st.expander('Methodology and limitations', expanded=False):
    st.markdown('''
1. Each **subject teacher enters a predicted final HSC mark** for the student on a 0–100 equivalent scale.
2. For each course and each year from **2021–2025**, the model estimates a **scaled mark per unit** by interpolating the HSC/scaled percentile pairs published in UAC Table A3.
3. The central scaled mark is the **mean of the available annual estimates**.
4. The model selects the **best 2 units of English + best 8 remaining units**.
5. The 10-unit scaled aggregate is converted to an ATAR using the mean of the **2021–2025 UAC aggregate→ATAR relationships**.
6. The report also shows each historical-year scenario, a teacher-prediction sensitivity range, and a Graham-style **ATAR contribution** for each subject.

**Important:** UAC states that Table A3 is not a one-to-one HSC→scaled conversion table. The same reported HSC mark can correspond to a range of underlying raw/scaled marks. This tool is therefore an **estimate for planning**, not an official UAC calculation.

For a 1-unit extension course, enter the prediction on a percentage-equivalent scale: **45/50 = 90**.
''')

# ---------- Student metadata ----------
st.subheader('1. Student details')
meta1, meta2, meta3, meta4 = st.columns([2.2, 1, 1.3, 1.3])
with meta1:
    student_name = st.text_input('Student name', placeholder='e.g. Alex Student')
with meta2:
    year_group = st.selectbox('Year group', ['Year 12', 'Year 11', 'Other'], index=0)
with meta3:
    student_id = st.text_input('Student ID (optional)', placeholder='Optional')
with meta4:
    prediction_date = st.date_input('Prediction date', value=date.today())

# ---------- Subject-teacher predictions ----------
courses = available_courses()
default = pd.DataFrame([
    {'Course': 'English Advanced', 'Subject teacher': '', 'Predicted HSC': 85.0},
    {'Course': 'Mathematics Advanced', 'Subject teacher': '', 'Predicted HSC': 85.0},
    {'Course': 'Chemistry', 'Subject teacher': '', 'Predicted HSC': 85.0},
    {'Course': 'Biology', 'Subject teacher': '', 'Predicted HSC': 85.0},
    {'Course': 'Modern History', 'Subject teacher': '', 'Predicted HSC': 85.0},
    {'Course': 'Business Studies', 'Subject teacher': '', 'Predicted HSC': 85.0},
])

st.subheader('2. Subject-teacher predictions')
st.caption('Each subject teacher enters their predicted final HSC course mark. Add or remove rows to match the student’s program.')

edited = st.data_editor(
    default,
    num_rows='dynamic',
    hide_index=True,
    use_container_width=True,
    column_config={
        'Course': st.column_config.SelectboxColumn('Course', options=courses, required=True, width='large'),
        'Subject teacher': st.column_config.TextColumn('Subject teacher', width='medium', help='Teacher responsible for the prediction.'),
        'Predicted HSC': st.column_config.NumberColumn('Predicted HSC', min_value=0.0, max_value=100.0, step=0.5, format='%.1f', required=True),
    },
    key='prediction_editor',
)

sensitivity = st.slider(
    'Teacher-prediction sensitivity (± HSC marks)',
    0.0, 5.0, 2.0, 0.5,
    help='Shows how the overall estimate moves if every subject prediction is this many HSC marks lower or higher.',
)

if st.button('Calculate estimate', type='primary', use_container_width=True):
    try:
        rows = [
            {
                'course': r['Course'],
                'teacher': r.get('Subject teacher', '') if pd.notna(r.get('Subject teacher', '')) else '',
                'mark': r['Predicted HSC'],
            }
            for _, r in edited.dropna(subset=['Course', 'Predicted HSC']).iterrows()
        ]
        result = estimate(rows, sensitivity)
        st.session_state['result'] = result
        st.session_state['input_rows'] = rows
        st.session_state['student_meta'] = {
            'student_name': student_name,
            'year_group': year_group,
            'student_id': student_id,
            'prediction_date': prediction_date.isoformat(),
        }
    except Exception as e:
        st.error(str(e))

result = st.session_state.get('result')
if result:
    input_rows = st.session_state.get('input_rows', [])
    saved_meta = st.session_state.get('student_meta', {})
    central = result['central_atar']
    year_atars = [x['atar'] for x in result['yearly'] if x['atar'] is not None]
    hist_range = f"{min(year_atars):.2f}–{max(year_atars):.2f}" if year_atars else 'Insufficient common-year data'
    sens_low = result['low']['atar']
    sens_high = result['high']['atar']
    sens_range = f"{sens_low:.2f}–{sens_high:.2f}" if sens_low is not None and sens_high is not None else 'Extends below published ATAR 50 table'

    st.divider()
    st.subheader('3. ATAR estimate')
    a, b, c, d = st.columns(4)
    a.metric('Central ATAR estimate', f'{central:.2f}' if central is not None else '<50')
    b.metric('Central 10-unit aggregate', f"{result['central_aggregate']:.1f}")
    c.metric('Historical scaling range', hist_range)
    d.metric(f'Sensitivity range (±{sensitivity:g})', sens_range)

    st.markdown('<div class="rh-note"><b>Interpretation:</b> the central estimate uses the mean historical scaled-mark estimate for each course. The historical range shows how the same teacher predictions would have translated under the available 2021–2025 scaling relationships.</div>', unsafe_allow_html=True)

    st.subheader('Subject contributions')
    sdf = pd.DataFrame(result['subjects'])
    sdf['Subject teacher'] = sdf['course'].map({r['course']: r.get('teacher', '') for r in input_rows}).fillna('')
    sdf['Counted units'] = sdf['course'].map(result['counts']).fillna(0).astype(int)
    sdf['ATAR contribution'] = sdf['atar_contribution'].apply(lambda x: '<50' if pd.isna(x) else f'{x:.2f}')
    sdf['Scaled mean /50'] = sdf['mean_scaled_per_unit'].round(2)
    sdf['Scaled range /50'] = sdf.apply(lambda r: f"{r['scaled_min']:.2f}–{r['scaled_max']:.2f}", axis=1)
    show = sdf[['course', 'Subject teacher', 'predicted_hsc', 'units', 'Counted units', 'Scaled mean /50', 'Scaled range /50', 'years_available', 'ATAR contribution']].copy()
    show.columns = ['Course', 'Subject teacher', 'Predicted HSC', 'Units', 'Counted units', 'Estimated scaled /50', '5-year scaled range', 'Years of data', 'ATAR contribution']
    st.dataframe(show, hide_index=True, use_container_width=True)
    st.caption('ATAR contribution is a diagnostic ATAR-equivalent level for that subject’s estimated scaled performance. It is not a number added directly to the final ATAR.')

    left, right = st.columns(2)
    with left:
        st.subheader('Historical-year scenarios')
        if result['yearly']:
            ydf = pd.DataFrame(result['yearly'])
            ydf['aggregate'] = ydf['aggregate'].round(1)
            ydf['atar'] = ydf['atar'].apply(lambda x: '<50' if pd.isna(x) else f'{x:.2f}')
            ydf.columns = ['Year', 'Aggregate', 'Estimated ATAR']
            st.dataframe(ydf, hide_index=True, use_container_width=True)
        else:
            st.info('No single historical year contains scaling data for every selected course. The central estimate still uses the available years for each subject.')

    with right:
        st.subheader('Which units counted?')
        counted = pd.DataFrame(result['chosen_units'])[['course', 'unit_index', 'scaled', 'english']]
        counted['scaled'] = counted['scaled'].round(2)
        counted['english'] = counted['english'].map({True: 'Yes', False: 'No'})
        counted.columns = ['Course', 'Unit within course', 'Scaled mark /50', 'English unit']
        st.dataframe(counted, hide_index=True, use_container_width=True)

    with st.expander('Annual scaled-mark estimates by subject'):
        detail = sdf[['course', 'Subject teacher', 'predicted_hsc'] + [f'scaled_{y}' for y in result['years']]].copy()
        detail.columns = ['Course', 'Subject teacher', 'Predicted HSC'] + [str(y) for y in result['years']]
        st.dataframe(detail.round(2), hide_index=True, use_container_width=True)

    # ---------- Excel export ----------
    st.subheader('4. Export report')
    try:
        xlsx_bytes = build_excel_export(saved_meta, input_rows, result, logo_path=LOGO_PATH)
        safe_name = re.sub(r'[^A-Za-z0-9_-]+', '_', saved_meta.get('student_name', '').strip()).strip('_') or 'Student'
        filename = f"{safe_name}_ATAR_Estimate_{saved_meta.get('prediction_date', date.today().isoformat())}.xlsx"
        st.download_button(
            'Download Excel report',
            data=xlsx_bytes,
            file_name=filename,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
            type='primary',
        )
        st.caption('The workbook contains Summary, Scaling Detail, ATAR Scenarios, Counted Units and Methodology sheets, including the subject teacher attached to each prediction and Reddam House North Shore branding.')
    except Exception as e:
        st.warning(f'Excel export could not be generated: {e}')

st.divider()
st.caption('Reddam House North Shore internal planning tool. Data source: Universities Admissions Centre (UAC), Reports on the Scaling of the NSW HSC 2021–2025. This estimator is independent and is not affiliated with UAC or NESA.')
