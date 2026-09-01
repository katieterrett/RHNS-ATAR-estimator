from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import hashlib
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

# Teacher prediction categories. Teachers select a broad HSC performance category;
# the estimator converts it to the representative HSC mark below before applying
# the existing five-year UAC scaling model.
BAND_TO_HSC_MARK = {
    '6A': 98.0,
    '6B': 94.0,
    '6C': 91.0,
    '5A': 88.0,
    '5B': 84.0,
    '5C': 81.0,
    '4A': 78.0,
    '4B': 74.0,
    '4C': 71.0,
    '3A': 68.0,
    '3B': 64.0,
    '3C': 61.0,
    '2A': 58.0,
    '2B': 54.0,
    '2C': 51.0,
}
BAND_OPTIONS = list(BAND_TO_HSC_MARK)

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
1. Each subject teacher selects the **predicted HSC performance category** that best represents the student's most likely final HSC course result (for example, 6A, 6B or 6C). The app converts that category to a representative HSC mark before scaling.
2. For each course and each year from **2021–2025**, the model estimates a **scaled mark per unit** by interpolating the HSC/scaled percentile pairs published in UAC Table A3.
3. The central scaled mark is the **mean of the available annual estimates**.
4. The model selects the **best 2 units of English + best 8 remaining units**.
5. The 10-unit scaled aggregate is converted to an ATAR using the mean of the **2021–2025 UAC aggregate→ATAR relationships**.
6. The report also shows each historical-year scenario, a teacher-prediction sensitivity range, and a Graham-style **ATAR contribution** for each subject.

**Important:** UAC states that Table A3 is not a one-to-one HSC→scaled conversion table. The same reported HSC mark can correspond to a range of underlying raw/scaled marks. This tool is therefore an **estimate for planning**, not an official UAC calculation.

For 1-unit Extension courses, use the same predicted-band categories. The representative mark is treated on the same 0–100 equivalent scale used by the estimator.
''')


# ---------------- Bulk import helpers ----------------
def _norm(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    return re.sub(r'[^a-z0-9]+', ' ', str(value).lower()).strip()


def _clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ''
    return str(value).strip()


def _clean_id(value) -> str:
    if value is None or pd.isna(value):
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return re.sub(r'\.0$', '', text)


def _as_mark(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.strip().replace('%', '')
        if not value:
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


STUDENT_ALIASES = {
    'student', 'student name', 'name', 'pupil', 'learner', 'student full name', 'full name'
}
ID_ALIASES = {
    'student id', 'student number', 'student no', 'student code', 'id', 'sid'
}
MARK_ALIASES = {
    'predicted hsc', 'predicted hsc mark', 'hsc predicted mark', 'hsc prediction',
    'predicted mark', 'prediction', 'predicted score', 'hsc predicted score',
    'teacher prediction', 'teacher predicted mark', 'mark', 'score'
}
COURSE_ALIASES = {
    'course', 'course name', 'subject', 'subject name', 'subject group', 'course group'
}
TEACHER_ALIASES = {
    'teacher', 'subject teacher', 'teacher name', 'class teacher'
}
YEAR_ALIASES = {
    'year group', 'year level', 'cohort', 'school year'
}


def _find_col(columns, aliases):
    for col in columns:
        if _norm(col) in aliases:
            return col
    return None


def _auto_course_match(label: str, courses: list[str]) -> str:
    """Conservative automatic matching of source subject labels to model course names."""
    if not label:
        return ''
    if label in courses:
        return label

    by_norm = {_norm(c): c for c in courses}
    n = _norm(label)
    if n in by_norm:
        return by_norm[n]

    # Common school naming variants.
    aliases = {
        'pdhpe': 'PDH&PE',
        'pdh pe': 'PDH&PE',
        'maths advanced': 'Mathematics Advanced',
        'math advanced': 'Mathematics Advanced',
        'maths extension 1': 'Mathematics Extension 1',
        'math extension 1': 'Mathematics Extension 1',
        'maths extension 2': 'Mathematics Extension 2',
        'math extension 2': 'Mathematics Extension 2',
        'maths standard 2': 'Mathematics Standard 2',
        'math standard 2': 'Mathematics Standard 2',
        'english eald': 'English EAL/D',
        'english eal d': 'English EAL/D',
        'society and culture': 'Society & Culture',
        'community and family studies': 'Community & Family Studies',
        'design and technology': 'Design & Technology',
        'textiles and design': 'Textiles & Design',
        'earth and environmental science': 'Earth & Environmental Science',
    }
    if n in aliases and aliases[n] in courses:
        return aliases[n]

    # Strip common class/group suffixes, e.g. "Chemistry HS (B)" or "Chemistry Year 12 A".
    simplified = re.sub(r'\([^)]*\)\s*$', '', str(label), flags=re.I)
    simplified = re.sub(r'\b(year|yr)\s*12\b', '', simplified, flags=re.I)
    simplified = re.sub(r'\bhsc\b|\bhs\b', '', simplified, flags=re.I)
    simplified = re.sub(r'\bclass\s*[a-z0-9-]+\b', '', simplified, flags=re.I)
    simplified = re.sub(r'\s+', ' ', simplified).strip(' -_')
    sn = _norm(simplified)
    if sn in by_norm:
        return by_norm[sn]
    if sn in aliases and aliases[sn] in courses:
        return aliases[sn]
    return ''


def _detect_header_row(book_bytes: bytes, sheet_name: str, courses: list[str]) -> int:
    raw = pd.read_excel(BytesIO(book_bytes), sheet_name=sheet_name, header=None, engine='openpyxl', nrows=15)
    course_norms = {_norm(c) for c in courses}
    for idx, row in raw.iterrows():
        labels = [_norm(v) for v in row.tolist() if pd.notna(v)]
        if not labels:
            continue
        has_student = any(v in STUDENT_ALIASES or v in ID_ALIASES for v in labels)
        has_mark = any(v in MARK_ALIASES for v in labels)
        has_course = any(v in COURSE_ALIASES or v in course_norms for v in labels)
        if has_student and (has_mark or has_course):
            return int(idx)
    return 0


def parse_bulk_workbook(book_bytes: bytes, courses: list[str]) -> tuple[pd.DataFrame, list[str]]:
    """Parse long, subject-sheet, or wide-format prediction workbooks.

    Returns a normalised dataframe with source subject labels still intact, plus parsing notes.
    """
    excel = pd.ExcelFile(BytesIO(book_bytes), engine='openpyxl')
    rows: list[dict] = []
    notes: list[str] = []

    for sheet in excel.sheet_names:
        try:
            header_row = _detect_header_row(book_bytes, sheet, courses)
            df = pd.read_excel(BytesIO(book_bytes), sheet_name=sheet, header=header_row, engine='openpyxl')
        except Exception as exc:
            notes.append(f'{sheet}: could not be read ({exc}).')
            continue

        df = df.dropna(axis=0, how='all').dropna(axis=1, how='all')
        if df.empty:
            continue

        student_col = _find_col(df.columns, STUDENT_ALIASES)
        id_col = _find_col(df.columns, ID_ALIASES)
        mark_col = _find_col(df.columns, MARK_ALIASES)
        course_col = _find_col(df.columns, COURSE_ALIASES)
        teacher_col = _find_col(df.columns, TEACHER_ALIASES)
        year_col = _find_col(df.columns, YEAR_ALIASES)

        if not student_col and not id_col:
            notes.append(f'{sheet}: skipped because no Student/Student ID column was found.')
            continue

        # Long format or one subject/group per worksheet.
        if mark_col:
            for _, r in df.iterrows():
                student_name = _clean_text(r.get(student_col)) if student_col else ''
                student_id = _clean_id(r.get(id_col)) if id_col else ''
                mark = _as_mark(r.get(mark_col))
                if (not student_name and not student_id) or mark is None:
                    continue
                source_course = _clean_text(r.get(course_col)) if course_col else str(sheet).strip()
                if not source_course:
                    continue
                rows.append({
                    'Student': student_name,
                    'Student ID': student_id,
                    'Source subject/group': source_course,
                    'Predicted HSC': mark,
                    'Subject teacher': _clean_text(r.get(teacher_col)) if teacher_col else '',
                    'Year group': _clean_text(r.get(year_col)) if year_col else '',
                    'Source sheet': sheet,
                })
            continue

        # Wide format: student metadata columns followed by subject columns.
        excluded = {c for c in [student_col, id_col, teacher_col, year_col, course_col] if c is not None}
        subject_columns = []
        for col in df.columns:
            if col in excluded:
                continue
            numeric = df[col].map(_as_mark)
            if numeric.notna().any():
                subject_columns.append(col)

        if not subject_columns:
            notes.append(f'{sheet}: no predicted-mark column or numeric subject columns were found.')
            continue

        for _, r in df.iterrows():
            student_name = _clean_text(r.get(student_col)) if student_col else ''
            student_id = _clean_id(r.get(id_col)) if id_col else ''
            if not student_name and not student_id:
                continue
            for col in subject_columns:
                mark = _as_mark(r.get(col))
                if mark is None:
                    continue
                rows.append({
                    'Student': student_name,
                    'Student ID': student_id,
                    'Source subject/group': str(col).strip(),
                    'Predicted HSC': mark,
                    'Subject teacher': '',
                    'Year group': _clean_text(r.get(year_col)) if year_col else '',
                    'Source sheet': sheet,
                })

    return pd.DataFrame(rows), notes


def _student_key(row) -> str:
    sid = _clean_id(row.get('Student ID'))
    if sid:
        return f'id:{_norm(sid)}'
    return f'name:{_norm(row.get("Student"))}'


def _build_student_reports(mapped: pd.DataFrame, default_year_group: str, prediction_date: date, sensitivity: float):
    work = mapped.copy()
    work['Student key'] = work.apply(_student_key, axis=1)

    # Remove exact duplicates but reject conflicting predictions for the same student/course.
    work = work.drop_duplicates(subset=['Student key', 'Course', 'Predicted HSC', 'Subject teacher'])
    duplicate_counts = work.groupby(['Student key', 'Course'], dropna=False).size()
    conflicts = set(duplicate_counts[duplicate_counts > 1].index.tolist())

    reports = []
    errors = []
    summary_rows = []

    for student_key, grp in work.groupby('Student key', sort=True):
        first = grp.iloc[0]
        student_name = _clean_text(first.get('Student')) or _clean_id(first.get('Student ID')) or 'Student'
        student_id = _clean_id(first.get('Student ID'))
        year_values = [x for x in grp['Year group'].map(_clean_text).tolist() if x]
        year_group = year_values[0] if year_values else default_year_group

        conflict_courses = [course for key, course in conflicts if key == student_key]
        if conflict_courses:
            msg = 'Conflicting duplicate predictions for: ' + ', '.join(sorted(conflict_courses))
            errors.append({'Student': student_name, 'Student ID': student_id, 'Issue': msg})
            continue

        rows = [
            {
                'course': r['Course'],
                'teacher': _clean_text(r.get('Subject teacher')),
                'mark': float(r['Predicted HSC']),
            }
            for _, r in grp.iterrows()
        ]
        try:
            result = estimate(rows, sensitivity)
            meta = {
                'student_name': student_name,
                'year_group': year_group,
                'student_id': student_id,
                'prediction_date': prediction_date.isoformat(),
            }
            reports.append({'student_meta': meta, 'input_rows': rows, 'result': result})
            summary_rows.append({
                'Student': student_name,
                'Student ID': student_id,
                'Subjects': len(rows),
                'Units entered': sum(int(s.get('units', 0)) for s in result.get('subjects', [])),
                'Central ATAR': '<50' if result.get('central_atar') is None else f"{result['central_atar']:.2f}",
                '10-unit aggregate': f"{result['central_aggregate']:.1f}",
            })
        except Exception as exc:
            errors.append({'Student': student_name, 'Student ID': student_id, 'Issue': str(exc)})

    return reports, pd.DataFrame(summary_rows), pd.DataFrame(errors)


# ---------------- Individual mode ----------------
def render_individual_mode(courses: list[str]):
    st.subheader('1. Student details')
    meta1, meta2, meta3, meta4 = st.columns([2.2, 1, 1.3, 1.3])
    with meta1:
        student_name = st.text_input('Student name', placeholder='e.g. Alex Student', key='individual_student_name')
    with meta2:
        year_group = st.selectbox('Year group', ['Year 12', 'Year 11', 'Other'], index=0, key='individual_year_group')
    with meta3:
        student_id = st.text_input('Student ID (optional)', placeholder='Optional', key='individual_student_id')
    with meta4:
        prediction_date = st.date_input('Prediction date', value=date.today(), key='individual_prediction_date')

    default = pd.DataFrame([
        {'Course': 'English Advanced', 'Predicted band': '5B'},
        {'Course': 'Mathematics Advanced', 'Predicted band': '5B'},
        {'Course': 'Chemistry', 'Predicted band': '5B'},
        {'Course': 'Biology', 'Predicted band': '5B'},
        {'Course': 'Modern History', 'Predicted band': '5B'},
        {'Course': 'Business Studies', 'Predicted band': '5B'},
    ])

    st.subheader('2. HSC predictions')
    st.caption(
        'For each course, select the category that best represents the student’s most likely final HSC course result. '
        'The app converts the category to a representative HSC mark automatically. Add or remove rows to match the student’s program.'
    )
    edited = st.data_editor(
        default,
        num_rows='dynamic',
        hide_index=True,
        use_container_width=True,
        column_config={
            'Course': st.column_config.SelectboxColumn('Course', options=courses, required=True, width='large'),
            'Predicted band': st.column_config.SelectboxColumn(
                'Predicted band',
                options=BAND_OPTIONS,
                required=True,
                width='medium',
                help='6A = high Band 6, 6B = mid Band 6, 6C = low Band 6, and likewise for Bands 5–2.'
            ),
        },
        key='prediction_editor',
    )

    with st.expander('Band-to-mark conversion used by the estimator', expanded=False):
        conversion = pd.DataFrame({
            'Predicted band': BAND_OPTIONS,
            'Representative HSC mark': [BAND_TO_HSC_MARK[b] for b in BAND_OPTIONS],
        })
        st.dataframe(conversion, hide_index=True, use_container_width=True)

    sensitivity = st.slider(
        'Teacher-prediction sensitivity (± HSC marks)',
        0.0, 5.0, 2.0, 0.5,
        help='Shows how the overall estimate moves if every subject prediction is this many HSC marks lower or higher.',
        key='individual_sensitivity',
    )

    if st.button('Calculate estimate', type='primary', use_container_width=True, key='individual_calculate'):
        try:
            rows = []
            for _, r in edited.dropna(subset=['Course', 'Predicted band']).iterrows():
                band = str(r['Predicted band']).strip()
                if band not in BAND_TO_HSC_MARK:
                    continue
                rows.append({
                    'course': r['Course'],
                    'teacher': '',
                    'band': band,
                    'mark': BAND_TO_HSC_MARK[band],
                })
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
        band_by_course = {r['course']: r.get('band', '') for r in input_rows}
        sdf['Predicted band'] = sdf['course'].map(band_by_course).fillna('')
        sdf['Counted units'] = sdf['course'].map(result['counts']).fillna(0).astype(int)
        sdf['ATAR contribution'] = sdf['atar_contribution'].apply(lambda x: '<50' if pd.isna(x) else f'{x:.2f}')
        sdf['Scaled mean /50'] = sdf['mean_scaled_per_unit'].round(2)
        sdf['Scaled range /50'] = sdf.apply(lambda r: f"{r['scaled_min']:.2f}–{r['scaled_max']:.2f}", axis=1)
        show = sdf[['course', 'Predicted band', 'predicted_hsc', 'units', 'Counted units', 'Scaled mean /50', 'Scaled range /50', 'years_available', 'ATAR contribution']].copy()
        show.columns = ['Course', 'Predicted band', 'Converted HSC mark', 'Units', 'Counted units', 'Estimated scaled /50', '5-year scaled range', 'Years of data', 'ATAR contribution']
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
            st.subheader('Courses included in ATAR aggregate')
            counted_units = pd.DataFrame(result['chosen_units'])
            if not counted_units.empty:
                counted = (
                    counted_units
                    .groupby('course', as_index=False)
                    .agg(
                        **{
                            'Units counted': ('unit_index', 'count'),
                            'Scaled mark /50 per unit': ('scaled', 'mean'),
                            'Aggregate contribution': ('scaled', 'sum'),
                            'English requirement': ('english', 'max'),
                        }
                    )
                )
                counted['Scaled mark /50 per unit'] = counted['Scaled mark /50 per unit'].round(2)
                counted['Aggregate contribution'] = counted['Aggregate contribution'].round(2)
                counted['English requirement'] = counted['English requirement'].map({True: 'Yes', False: ''})
                counted = counted.rename(columns={'course': 'Course'})
                counted = counted.sort_values(
                    by=['English requirement', 'Aggregate contribution'],
                    ascending=[False, False],
                    kind='stable',
                )
                st.dataframe(counted, hide_index=True, use_container_width=True)
                st.caption(
                    'Aggregate contribution is the scaled mark contributed by each course to the student’s 10-unit ATAR aggregate. '
                    'The contributions are added together, then the total aggregate is converted to an estimated ATAR using the historical UAC aggregate→ATAR relationship.'
                )
                st.metric('Total 10-unit aggregate', f"{result['central_aggregate']:.2f}")
            else:
                st.info('No counted units are available for this estimate.')

        with st.expander('Annual scaled-mark estimates by subject'):
            detail = sdf[['course', 'Predicted band', 'predicted_hsc'] + [f'scaled_{y}' for y in result['years']]].copy()
            detail.columns = ['Course', 'Predicted band', 'Converted HSC mark'] + [str(y) for y in result['years']]
            st.dataframe(detail.round(2), hide_index=True, use_container_width=True)

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
                key='individual_download',
            )
            st.caption('The workbook contains Summary, Scaling Detail, ATAR Scenarios, Counted Units and Methodology sheets with Reddam House North Shore branding.')
        except Exception as e:
            st.warning(f'Excel export could not be generated: {e}')


# ---------------- Bulk mode ----------------
def render_bulk_mode(courses: list[str]):
    st.subheader('1. Upload cohort predictions')
    st.markdown(
        '<div class="rh-note"><b>Bulk mode:</b> upload one Excel workbook containing teacher-predicted HSC marks. '
        'The app groups the records by student, calculates each student separately, and produces one Excel workbook with one Summary-style tab per student.</div>',
        unsafe_allow_html=True,
    )

    with st.expander('Supported Excel layouts', expanded=False):
        st.markdown('''
**Subject/group sheets** — one worksheet per subject or class group. Each sheet needs a **Student** (or Student ID) column and a **Predicted HSC**/Mark column. The sheet name is used as the subject label and can be mapped to the official course name after upload.

**Long list** — one or more worksheets with columns such as **Student**, **Student ID**, **Course/Subject**, **Predicted HSC**, and optionally **Subject teacher** and **Year group**.

**Wide table** — one row per student, with Student/Student ID columns followed by subject columns containing predicted HSC marks.

Predictions must be on a **0–100 equivalent scale**. For a 1-unit extension course, convert the mark to a percentage-equivalent value first (for example, 45/50 = 90).
''')

    uploaded = st.file_uploader('Upload Excel workbook (.xlsx)', type=['xlsx'], key='bulk_upload')
    if uploaded is None:
        return

    book_bytes = uploaded.getvalue()
    upload_sig = hashlib.sha1(book_bytes).hexdigest()
    if st.session_state.get('bulk_upload_sig') != upload_sig:
        st.session_state['bulk_upload_sig'] = upload_sig
        for key in ['bulk_reports', 'bulk_summary', 'bulk_errors', 'bulk_date']:
            st.session_state.pop(key, None)
    try:
        parsed, notes = parse_bulk_workbook(book_bytes, courses)
    except Exception as exc:
        st.error(f'The workbook could not be read: {exc}')
        return

    if parsed.empty:
        st.error('No usable prediction rows were found. Check the column headings and workbook layout.')
        if notes:
            with st.expander('Import notes'):
                for note in notes:
                    st.write(f'• {note}')
        return

    if notes:
        with st.expander('Import notes'):
            for note in notes:
                st.write(f'• {note}')

    # Validate marks before course mapping.
    invalid_marks = parsed[(parsed['Predicted HSC'] < 0) | (parsed['Predicted HSC'] > 100)]
    if not invalid_marks.empty:
        st.error('Some predicted HSC marks are outside the permitted 0–100 range.')
        st.dataframe(invalid_marks[['Student', 'Student ID', 'Source subject/group', 'Predicted HSC', 'Source sheet']], hide_index=True, use_container_width=True)
        return

    st.subheader('2. Check subject matching')
    source_labels = sorted(parsed['Source subject/group'].dropna().astype(str).unique(), key=str.casefold)
    mapping_default = pd.DataFrame({
        'Subject/group in upload': source_labels,
        'ATAR course': [_auto_course_match(label, courses) for label in source_labels],
    })
    mapping_edited = st.data_editor(
        mapping_default,
        hide_index=True,
        use_container_width=True,
        disabled=['Subject/group in upload'],
        column_config={
            'Subject/group in upload': st.column_config.TextColumn('Subject/group in upload', width='large'),
            'ATAR course': st.column_config.SelectboxColumn('ATAR course', options=[''] + courses, width='large'),
        },
        key=f'bulk_course_mapping_{upload_sig[:10]}',
    )

    mapping = dict(zip(mapping_edited['Subject/group in upload'], mapping_edited['ATAR course']))
    mapped = parsed.copy()
    mapped['Course'] = mapped['Source subject/group'].map(mapping).fillna('')

    unmapped = sorted(mapped.loc[mapped['Course'].eq(''), 'Source subject/group'].unique().tolist(), key=str.casefold)
    if unmapped:
        st.warning('Map every detected subject/group before calculating: ' + ', '.join(unmapped))

    preview = mapped[['Student', 'Student ID', 'Course', 'Predicted HSC', 'Subject teacher', 'Source sheet']].copy()
    st.caption('Preview of records after subject matching')
    st.dataframe(preview, hide_index=True, use_container_width=True, height=min(520, 38 + 35 * min(len(preview), 14)))

    unique_students = mapped.apply(_student_key, axis=1).nunique()
    m1, m2, m3 = st.columns(3)
    m1.metric('Students detected', unique_students)
    m2.metric('Prediction rows', len(mapped))
    m3.metric('Subject/groups detected', len(source_labels))

    st.subheader('3. Cohort settings')
    c1, c2, c3 = st.columns([1.2, 1.3, 2.0])
    with c1:
        default_year_group = st.selectbox('Default year group', ['Year 12', 'Year 11', 'Other'], index=0, key='bulk_year_group')
    with c2:
        prediction_date = st.date_input('Prediction date', value=date.today(), key='bulk_prediction_date')
    with c3:
        sensitivity = st.slider(
            'Teacher-prediction sensitivity (± HSC marks)',
            0.0, 5.0, 2.0, 0.5,
            key='bulk_sensitivity',
            help='Applied consistently to every student in the cohort.',
        )

    if st.button('Calculate cohort estimates', type='primary', use_container_width=True, disabled=bool(unmapped), key='bulk_calculate'):
        reports, summary, errors = _build_student_reports(mapped, default_year_group, prediction_date, sensitivity)
        st.session_state['bulk_reports'] = reports
        st.session_state['bulk_summary'] = summary
        st.session_state['bulk_errors'] = errors
        st.session_state['bulk_date'] = prediction_date.isoformat()

    reports = st.session_state.get('bulk_reports', [])
    summary = st.session_state.get('bulk_summary', pd.DataFrame())
    errors = st.session_state.get('bulk_errors', pd.DataFrame())

    if reports or (isinstance(errors, pd.DataFrame) and not errors.empty):
        st.divider()
        st.subheader('4. Cohort results')
        r1, r2 = st.columns(2)
        r1.metric('Student reports ready', len(reports))
        r2.metric('Students requiring attention', 0 if not isinstance(errors, pd.DataFrame) else len(errors))

        if isinstance(summary, pd.DataFrame) and not summary.empty:
            st.dataframe(summary, hide_index=True, use_container_width=True)

        if isinstance(errors, pd.DataFrame) and not errors.empty:
            st.warning('The following students were not included in the export. Correct their source data and re-run the cohort calculation.')
            st.dataframe(errors, hide_index=True, use_container_width=True)

        if reports:
            try:
                xlsx_bytes = build_bulk_excel_export(reports, logo_path=LOGO_PATH)
                report_date = st.session_state.get('bulk_date', date.today().isoformat())
                st.download_button(
                    'Download cohort Excel report',
                    data=xlsx_bytes,
                    file_name=f'RHNS_Cohort_ATAR_Estimates_{report_date}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                    type='primary',
                    key='bulk_download',
                )
                st.caption('The downloaded workbook contains one Summary-style worksheet per successfully calculated student and no calculation-detail tabs.')
            except Exception as exc:
                st.warning(f'Cohort Excel export could not be generated: {exc}')


courses = available_courses()
render_individual_mode(courses)

st.divider()
st.caption('Reddam House North Shore internal planning tool. Data source: Universities Admissions Centre (UAC), Reports on the Scaling of the NSW HSC 2021–2025. This estimator is independent and is not affiliated with UAC or NESA.')
