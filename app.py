from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import hashlib
import re

import pandas as pd
import streamlit as st

from model import estimate, available_courses
from export_excel import build_excel_export, build_bulk_excel_export

SCHOOL_NAME = 'Reddam House North Shore'
NAVY = '#1B293B'
GOLD = '#A68E3A'
PALE = '#F6F3EB'
LOGO_PATH = Path(__file__).resolve().parent / 'assets' / 'reddam_house_north_shore_logo.png'

# Teacher prediction categories. Standard 2-unit courses use 6A/6B/6C etc.
# Extension courses use E4A/E4B/E4C etc. All values are converted to the
# 0–100 equivalent scale expected by the existing five-year UAC model.
STANDARD_BAND_TO_HSC_MARK = {
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

# Extension performance bands are reported out of 50. These A/B/C categories
# subdivide each E-band into high / mid / low regions, then convert the
# representative mark to a 0–100 equivalent value for the estimator.
EXTENSION_BAND_TO_HSC_MARK = {
    'E4A': 98.0,  # about 49/50
    'E4B': 95.0,  # about 47.5/50
    'E4C': 91.0,  # about 45.5/50
    'E3A': 86.0,  # about 43/50
    'E3B': 80.0,  # about 40/50
    'E3C': 73.0,  # about 36.5/50
    'E2A': 66.0,  # about 33/50
    'E2B': 60.0,  # about 30/50
    'E2C': 53.0,  # about 26.5/50
    'E1A': 41.0,  # about 20.5/50
    'E1B': 25.0,  # about 12.5/50
    'E1C': 8.0,   # about 4/50
}

STANDARD_BAND_OPTIONS = list(STANDARD_BAND_TO_HSC_MARK)
EXTENSION_BAND_OPTIONS = list(EXTENSION_BAND_TO_HSC_MARK)
BAND_OPTIONS = STANDARD_BAND_OPTIONS + EXTENSION_BAND_OPTIONS
BAND_TO_HSC_MARK = {**STANDARD_BAND_TO_HSC_MARK, **EXTENSION_BAND_TO_HSC_MARK}

HMS_COURSE = 'Health and Movement Science'
HMS_PROXY_COURSE = 'PDH&PE'

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
7. **Health and Movement Science (HMS)** is treated as a special provisional case: until genuine HMS historical scaling data are available in this model, the app uses the **2021–2025 PDH&PE UAC scaling curves as a proxy**. HMS is still shown as HMS in all student outputs.

**Important:** UAC states that Table A3 is not a one-to-one HSC→scaled conversion table. The same reported HSC mark can correspond to a range of underlying raw/scaled marks. This tool is therefore an **estimate for planning**, not an official UAC calculation.

For 1-unit Extension courses, use **E4A/E4B/E4C, E3A/E3B/E3C, E2A/E2B/E2C or E1A/E1B/E1C**. The app converts these Extension predictions to the same 0–100 equivalent scale used internally by the estimator.

**HMS proxy:** Health and Movement Science is currently scaled using the historical **PDH&PE** UAC relationship as a provisional proxy. This assumption is clearly identified in the on-screen and Excel outputs and should be replaced when genuine HMS scaling data are available.
''')


# ---------------- RHNS cohort workbook helpers ----------------
SURNAME_ALIASES = {'surname', 'last name', 'family name'}
FORENAME_ALIASES = {'preferred forename', 'preferred name', 'forename', 'first name', 'given name'}
BAND_ALIASES = {'hsc band prediction', 'hsc band prediction code', 'predicted band', 'hsc prediction band'}
BULK_SENSITIVITY = 2.0


def _norm(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    return re.sub(r'[^a-z0-9]+', ' ', str(value).casefold()).strip()


def _clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ''
    return re.sub(r'\s+', ' ', str(value)).strip()


def _normalise_band(value) -> str:
    if value is None or pd.isna(value):
        return ''
    return re.sub(r'\s+', '', str(value)).upper().strip()


def _is_extension_course(course: str) -> bool:
    return 'extension' in str(course or '').casefold()


def _valid_prediction_code(course: str, band: str) -> bool:
    band = _normalise_band(band)
    if _is_extension_course(course):
        return band in EXTENSION_BAND_TO_HSC_MARK
    return band in STANDARD_BAND_TO_HSC_MARK


def _prediction_mark(course: str, band: str) -> float:
    band = _normalise_band(band)
    if not _valid_prediction_code(course, band):
        if _is_extension_course(course):
            raise ValueError(
                f'{course} is an Extension course. Use an Extension prediction code such as '
                'E4A, E4B, E4C, E3A, E3B or E3C.'
            )
        raise ValueError(
            f'{course} is not an Extension course. Use a standard prediction code such as '
            '6A, 6B, 6C, 5A, 5B or 5C.'
        )
    return BAND_TO_HSC_MARK[band]


def _student_key(surname: str, forename: str) -> str:
    return f'{_norm(surname)}|{_norm(forename)}'


def _find_col(columns, aliases):
    for col in columns:
        if _norm(col) in aliases:
            return col
    return None


def _infer_rhns_course(sheet_name: str, class_title: str, courses: list[str]) -> str:
    """Map the class names in the supplied RHNS workbook to UAC course names.

    Blank means the course is intentionally left unmapped and will be excluded unless
    the user chooses a course in the mapping table.
    """
    s = _norm(sheet_name)
    t = _norm(class_title)

    # Two French class codes are more informative than the visible class title.
    if 'frbhs' in s and 'French Beginners' in courses:
        return 'French Beginners'
    if 'frchs' in s and 'French Continuers' in courses:
        return 'French Continuers'

    # Health and Movement Science is included using PDH&PE as an explicit
    # provisional scaling proxy in model.py. Keep the displayed course name as HMS.
    if 'health and movement science' in t and HMS_COURSE in courses:
        return HMS_COURSE

    # Match Extension courses whose official UAC course name appears in the class title.
    # Specific school aliases such as "Maths Extension 1" are handled below.
    for course in courses:
        if 'Extension' in course and _norm(course) in t:
            return course

    exact_hints = [
        ('english extension 2', 'English Extension 2'),
        ('english extension 1', 'English Extension 1'),
        ('maths extension 2', 'Mathematics Extension 2'),
        ('maths extension 1', 'Mathematics Extension 1'),
        ('english advanced', 'English Advanced'),
        ('english standard', 'English Standard'),
        ('english eald', 'English EAL/D'),
        # Accept both RHNS wording ("Maths Advanced") and the official
        # NESA/UAC wording ("Mathematics Advanced"). This is needed for
        # the accelerated Year 12 Mathematics Advanced sheet.
        ('mathematics advanced', 'Mathematics Advanced'),
        ('maths advanced', 'Mathematics Advanced'),
        ('mathematics standard', 'Mathematics Standard 2'),
        ('maths standard', 'Mathematics Standard 2'),
        ('ancient history', 'Ancient History'),
        ('modern history', 'Modern History'),
        ('health and movement science', HMS_COURSE),
        ('business studies', 'Business Studies'),
        ('chinese continuers', 'Chinese Continuers'),
        ('chinese in context', 'Chinese in Context'),
        ('software engineering', 'Software Engineering'),
        ('legal studies', 'Legal Studies'),
        ('visual arts', 'Visual Arts'),
        ('biology', 'Biology'),
        ('chemistry', 'Chemistry'),
        ('economics', 'Economics'),
        ('geography', 'Geography'),
        ('physics', 'Physics'),
        ('dance', 'Dance'),
        ('drama', 'Drama'),
        ('music 1', 'Music 1'),
        ('music 2', 'Music 2'),
    ]
    for phrase, course in exact_hints:
        if phrase in t and course in courses:
            return course
    return ''


def _find_rhns_header_row(raw: pd.DataFrame) -> int | None:
    for idx, row in raw.head(12).iterrows():
        labels = {_norm(v) for v in row.tolist() if not pd.isna(v)}
        if labels & SURNAME_ALIASES and labels & FORENAME_ALIASES and labels & BAND_ALIASES:
            return int(idx)
    return None


def parse_rhns_band_workbook(book_bytes: bytes, courses: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Read the supplied RHNS subject-tab workbook.

    Expected layout per subject tab:
      row 1: class title
      row 2: Surname | Preferred Forename | HSC Band Prediction
      remaining rows: students
    The parser also tolerates the header appearing a few rows later.
    """
    excel = pd.ExcelFile(BytesIO(book_bytes), engine='openpyxl')
    rows: list[dict] = []
    groups: list[dict] = []
    notes: list[str] = []

    for sheet in excel.sheet_names:
        try:
            raw = pd.read_excel(BytesIO(book_bytes), sheet_name=sheet, header=None, engine='openpyxl', nrows=15)
            header_row = _find_rhns_header_row(raw)
            if header_row is None:
                notes.append(f'{sheet}: skipped because the Surname / Preferred Forename / HSC Band Prediction header was not found.')
                continue

            class_title = ''
            for r in range(header_row):
                for value in raw.iloc[r].tolist():
                    text = _clean_text(value)
                    if text:
                        class_title = text
                        break
                if class_title:
                    break
            class_title = class_title or sheet

            df = pd.read_excel(BytesIO(book_bytes), sheet_name=sheet, header=header_row, engine='openpyxl')
            df = df.dropna(axis=0, how='all')
            surname_col = _find_col(df.columns, SURNAME_ALIASES)
            forename_col = _find_col(df.columns, FORENAME_ALIASES)
            band_col = _find_col(df.columns, BAND_ALIASES)
            if surname_col is None or forename_col is None or band_col is None:
                notes.append(f'{sheet}: required columns could not be identified.')
                continue

            default_course = _infer_rhns_course(sheet, class_title, courses)
            groups.append({
                'Source sheet': sheet,
                'Class title': class_title,
                'ATAR course': default_course,
            })

            for _, r in df.iterrows():
                surname = _clean_text(r.get(surname_col))
                forename = _clean_text(r.get(forename_col))
                if not surname and not forename:
                    continue
                band = _normalise_band(r.get(band_col))
                rows.append({
                    'Student key': _student_key(surname, forename),
                    'Surname': surname,
                    'Preferred Forename': forename,
                    'Student': f'{forename} {surname}'.strip(),
                    'Predicted band': band,
                    'Source sheet': sheet,
                    'Class title': class_title,
                })
        except Exception as exc:
            notes.append(f'{sheet}: could not be read ({exc}).')

    return pd.DataFrame(rows), pd.DataFrame(groups), notes


def _atar_range_text(result: dict) -> str:
    central = result.get('central_atar')
    if central is None:
        return '<50'
    low = max(0.0, float(central) - 5.0)
    high = min(99.95, float(central) + 5.0)
    return f'{low:.2f}–{high:.2f}'


def _build_rhns_student_reports(records: pd.DataFrame, prediction_date: date, year_group: str = 'Year 11'):
    """Group rows by normalised surname+forename and run the estimator for each student."""
    reports: list[dict] = []
    summary_rows: list[dict] = []
    errors: list[dict] = []

    if records.empty:
        return reports, pd.DataFrame(summary_rows), pd.DataFrame(errors)

    for student_key, grp in records.groupby('Student key', sort=True):
        first = grp.iloc[0]
        surname = _clean_text(first.get('Surname'))
        forename = _clean_text(first.get('Preferred Forename'))
        student_name = f'{forename} {surname}'.strip() or 'Student'

        # Resolve each course separately. If the workbook contains a duplicate row/sheet
        # for the same student and course, a single valid prediction takes precedence over
        # blank/invalid duplicate values. This is important for the current RHNS workbook,
        # which contains an old duplicate Mathematics Advanced sheet with placeholder values.
        # A student is only blocked when a course has no valid prediction, or when two
        # different valid predictions exist for the same course.
        course_rows = []
        course_issues = []
        for course, cgrp in grp.groupby('Course', sort=False):
            course = str(course)
            valid_bands = sorted({
                _normalise_band(v)
                for v in cgrp['Predicted band'].tolist()
                if _valid_prediction_code(course, v)
            })

            if len(valid_bands) > 1:
                course_issues.append(
                    f'{course} has conflicting valid predictions: {", ".join(valid_bands)}'
                )
                continue

            if len(valid_bands) == 1:
                course_rows.append((course, valid_bands[0]))
                continue

            supplied = sorted({
                _normalise_band(v) or 'blank'
                for v in cgrp['Predicted band'].tolist()
            })
            expected = 'E4A–E1C' if _is_extension_course(course) else '6A–2C'
            course_issues.append(
                f'{course} has no valid prediction ({", ".join(supplied)}; expected {expected})'
            )

        if course_issues:
            errors.append({
                'Student': student_name,
                'Issue': '; '.join(course_issues),
            })
            continue

        input_rows = [
            {
                'course': course,
                'teacher': '',
                'band': band,
                'mark': _prediction_mark(course, band),
            }
            for course, band in course_rows
        ]

        try:
            result = estimate(input_rows, BULK_SENSITIVITY)
        except Exception as exc:
            errors.append({'Student': student_name, 'Issue': str(exc)})
            continue

        meta = {
            'student_name': student_name,
            'surname': surname,
            'forename': forename,
            'year_group': year_group,
            'student_id': '',
            'prediction_date': prediction_date.isoformat(),
        }
        reports.append({'student_meta': meta, 'input_rows': input_rows, 'result': result})
        summary_rows.append({
            'Student': student_name,
            'Subjects included': len(input_rows),
            'ATAR range': _atar_range_text(result),
        })

    reports.sort(key=lambda r: (r['student_meta']['surname'].casefold(), r['student_meta']['forename'].casefold()))
    summary = pd.DataFrame(summary_rows)
    if not summary.empty:
        summary = summary.sort_values('Student', key=lambda s: s.str.casefold()).reset_index(drop=True)
    return reports, summary, pd.DataFrame(errors)


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
                help=(
                    'Standard courses: 6A/6B/6C through 2A/2B/2C. '
                    'Extension courses: E4A/E4B/E4C through E1A/E1B/E1C.'
                )
            ),
        },
        key='prediction_editor',
    )

    with st.expander('Band-to-mark conversion used by the estimator', expanded=False):
        st.markdown('**Standard courses**')
        standard_conversion = pd.DataFrame({
            'Predicted band': STANDARD_BAND_OPTIONS,
            'Representative HSC mark /100': [STANDARD_BAND_TO_HSC_MARK[b] for b in STANDARD_BAND_OPTIONS],
        })
        st.dataframe(standard_conversion, hide_index=True, use_container_width=True)

        st.markdown('**Extension courses**')
        extension_conversion = pd.DataFrame({
            'Predicted band': EXTENSION_BAND_OPTIONS,
            'Representative Extension mark /50': [EXTENSION_BAND_TO_HSC_MARK[b] / 2 for b in EXTENSION_BAND_OPTIONS],
            'Model input /100 equivalent': [EXTENSION_BAND_TO_HSC_MARK[b] for b in EXTENSION_BAND_OPTIONS],
        })
        st.dataframe(extension_conversion, hide_index=True, use_container_width=True)

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
                course = str(r['Course']).strip()
                band = _normalise_band(r['Predicted band'])
                rows.append({
                    'course': course,
                    'teacher': '',
                    'band': band,
                    'mark': _prediction_mark(course, band),
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
        show.columns = ['Course', 'Predicted band', 'Converted HSC mark /100 equivalent', 'Units', 'Counted units', 'Estimated scaled /50', '5-year scaled range', 'Years of data', 'ATAR contribution']
        st.dataframe(show, hide_index=True, use_container_width=True)
        st.caption('ATAR contribution is a diagnostic ATAR-equivalent level for that subject’s estimated scaled performance. It is not a number added directly to the final ATAR.')
        if any(bool(x.get('provisional_proxy')) for x in result.get('subjects', [])):
            st.warning('Health and Movement Science is provisional: its scaled mark and ATAR contribution currently use 2021–2025 PDH&PE UAC scaling data as a proxy.')

        st.subheader('4. Supporting calculation detail')
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
            detail.columns = ['Course', 'Predicted band', 'Converted HSC mark /100 equivalent'] + [str(y) for y in result['years']]
            st.dataframe(detail.round(2), hide_index=True, use_container_width=True)

        st.subheader('5. Excel output')
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


# ---------------- Section 6: bulk cohort processing ----------------
def render_bulk_section(courses: list[str]):
    st.divider()
    st.subheader('6. Bulk cohort processing')
    st.markdown(
        '<div class="rh-note"><b>Cohort workflow:</b> upload the completed RHNS HSC band-prediction workbook. '
        'The app matches students across the subject tabs using surname + preferred forename, converts standard 6A/6B/6C-style predictions and '
        'Extension E4A/E4B/E4C-style predictions to representative HSC-equivalent marks, runs each student through the same five-year estimator, '
        'and creates one Excel workbook with one prediction-summary tab per student.</div>',
        unsafe_allow_html=True,
    )
    st.caption('Name matching is case-insensitive and ignores extra spaces. The app does not use fuzzy matching, so spelling differences are flagged rather than guessed.')

    uploaded = st.file_uploader(
        'Upload completed HSC band prediction workbook (.xlsx)',
        type=['xlsx'],
        key='rhns_bulk_band_upload',
    )
    if uploaded is None:
        return

    book_bytes = uploaded.getvalue()
    upload_sig = hashlib.sha1(book_bytes).hexdigest()
    if st.session_state.get('rhns_bulk_sig') != upload_sig:
        st.session_state['rhns_bulk_sig'] = upload_sig
        for key in ['rhns_bulk_reports', 'rhns_bulk_summary', 'rhns_bulk_errors', 'rhns_bulk_date']:
            st.session_state.pop(key, None)

    try:
        parsed, groups, notes = parse_rhns_band_workbook(book_bytes, courses)
    except Exception as exc:
        st.error(f'The workbook could not be read: {exc}')
        return

    if parsed.empty or groups.empty:
        st.error('No usable subject tabs were found in the workbook.')
        if notes:
            with st.expander('Import notes'):
                for note in notes:
                    st.write(f'• {note}')
        return

    if notes:
        with st.expander('Import notes'):
            for note in notes:
                st.write(f'• {note}')

    st.markdown('**Check subject matching**')
    st.caption('The app has pre-matched the class titles in this workbook to the UAC course names used by the estimator. Please check these before calculating.')
    mapping = st.data_editor(
        groups,
        hide_index=True,
        use_container_width=True,
        disabled=['Source sheet', 'Class title'],
        column_config={
            'Source sheet': st.column_config.TextColumn('Source tab', width='medium'),
            'Class title': st.column_config.TextColumn('Class title', width='large'),
            'ATAR course': st.column_config.SelectboxColumn('ATAR course', options=[''] + courses, width='large', help='Leave blank to exclude a course from this five-year model.'),
        },
        key=f'rhns_course_mapping_{upload_sig[:10]}',
    )

    course_by_sheet = dict(zip(mapping['Source sheet'], mapping['ATAR course']))
    mapped = parsed.copy()
    mapped['Course'] = mapped['Source sheet'].map(course_by_sheet).fillna('')

    excluded_groups = mapping[mapping['ATAR course'].fillna('').eq('')]
    if not excluded_groups.empty:
        st.warning(
            'The following class groups are currently excluded because they are not mapped to a course in the five-year model: '
            + '; '.join(excluded_groups['Class title'].astype(str).tolist())
            + '. Students can still be calculated if they have at least 10 other eligible units including 2 units of English.'
        )

    included = mapped[mapped['Course'].ne('')].copy()
    if (included['Course'] == HMS_COURSE).any():
        st.info('HMS proxy in use: Health and Movement Science predictions will use the 2021–2025 PDH&PE UAC scaling relationship as a provisional proxy. HMS will remain labelled as Health and Movement Science in the student reports.')
    nonblank = included[included['Predicted band'].ne('')]
    valid_mask = nonblank.apply(
        lambda r: _valid_prediction_code(r.get('Course', ''), r.get('Predicted band', '')),
        axis=1,
    )
    invalid = nonblank[~valid_mask]
    missing = included[included['Predicted band'].eq('')]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric('Students detected', parsed['Student key'].nunique())
    m2.metric('Subject tabs detected', len(groups))
    m3.metric('Predictions entered', len(nonblank) - len(invalid))
    m4.metric('Missing predictions', len(missing))

    if not invalid.empty:
        st.warning(
            'Some prediction codes do not match the mapped course type. '
            'The cohort can still be processed. If the same student/course also has one valid prediction elsewhere in the workbook, '
            'the valid prediction will be used; otherwise that student will be listed under Students requiring attention. '
            'Use 6A–2C for standard courses and E4A–E1C for Extension courses.'
        )
        invalid_show = invalid[['Student', 'Class title', 'Course', 'Predicted band']].copy()
        invalid_show['Expected code type'] = invalid_show['Course'].apply(
            lambda c: 'Extension: E4A–E1C' if _is_extension_course(c) else 'Standard: 6A–2C'
        )
        st.dataframe(
            invalid_show.rename(columns={'Class title': 'Subject/class'}),
            hide_index=True,
            use_container_width=True,
        )

    if not missing.empty:
        with st.expander(f'Missing predictions ({len(missing)})'):
            st.dataframe(
                missing[['Student', 'Class title']].rename(columns={'Class title': 'Subject/class'}),
                hide_index=True,
                use_container_width=True,
            )

    with st.expander('Preview matched student records', expanded=False):
        preview = included[['Student', 'Course', 'Predicted band', 'Source sheet']].sort_values(['Student', 'Course'])
        st.dataframe(preview, hide_index=True, use_container_width=True, height=500)

    c1, c2 = st.columns([1.2, 1.5])
    with c1:
        prediction_date = st.date_input('Cohort prediction date', value=date.today(), key='rhns_bulk_prediction_date')
    with c2:
        st.markdown('**Prediction sensitivity:** ±2 HSC marks')
        st.caption('The same ±2 sensitivity is applied on the 0–100 equivalent prediction scale for every student. For a 1-unit Extension course this is equivalent to ±1 mark out of 50. It does not change the central prediction.')

    can_calculate = not included.empty
    if st.button('Calculate cohort estimates', type='primary', use_container_width=True, disabled=not can_calculate, key='rhns_bulk_calculate'):
        reports, summary, errors = _build_rhns_student_reports(included, prediction_date, year_group='Year 11')
        st.session_state['rhns_bulk_reports'] = reports
        st.session_state['rhns_bulk_summary'] = summary
        st.session_state['rhns_bulk_errors'] = errors
        st.session_state['rhns_bulk_date'] = prediction_date.isoformat()

    reports = st.session_state.get('rhns_bulk_reports', [])
    summary = st.session_state.get('rhns_bulk_summary', pd.DataFrame())
    errors = st.session_state.get('rhns_bulk_errors', pd.DataFrame())

    if reports or (isinstance(errors, pd.DataFrame) and not errors.empty):
        st.markdown('**Cohort results**')
        r1, r2 = st.columns(2)
        r1.metric('Student reports ready', len(reports))
        r2.metric('Students requiring attention', 0 if not isinstance(errors, pd.DataFrame) else len(errors))

        if isinstance(summary, pd.DataFrame) and not summary.empty:
            st.dataframe(summary, hide_index=True, use_container_width=True)

        if isinstance(errors, pd.DataFrame) and not errors.empty:
            st.warning('These students were not included in the bulk workbook. The app has not guessed or silently filled missing data.')
            st.dataframe(errors, hide_index=True, use_container_width=True)

        if reports:
            try:
                xlsx_bytes = build_bulk_excel_export(reports, logo_path=LOGO_PATH)
                report_date = st.session_state.get('rhns_bulk_date', date.today().isoformat())
                st.download_button(
                    'Download cohort Excel workbook',
                    data=xlsx_bytes,
                    file_name=f'RHNS_2027_Cohort_ATAR_Prediction_Summaries_{report_date}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                    type='primary',
                    key='rhns_bulk_download',
                )
                st.caption('The workbook contains one Summary-style worksheet per successfully calculated student. It does not include the individual calculation-detail tabs.')
            except Exception as exc:
                st.warning(f'Cohort Excel export could not be generated: {exc}')


courses = available_courses()
render_individual_mode(courses)
render_bulk_section(courses)

st.divider()
st.caption('Reddam House North Shore internal planning tool. Data source: Universities Admissions Centre (UAC), Reports on the Scaling of the NSW HSC 2021–2025. This estimator is independent and is not affiliated with UAC or NESA.')
