from __future__ import annotations

from io import BytesIO
from datetime import datetime
from pathlib import Path
import math
import re
import xlsxwriter

NAVY = '#1B293B'
GOLD = '#A68E3A'
LIGHT_GOLD = '#F3EFE4'
PALE_BLUE = '#EAF0F6'
MID_GREY = '#6B7280'
LIGHT_GREY = '#F4F5F7'
DARK = '#1F2937'
WHITE = '#FFFFFF'
GREEN = '#E8F4EC'

HMS_COURSE = 'Health and Movement Science'
HMS_PROXY_COURSE = 'PDH&PE'


def _safe_num(value):
    if value is None:
        return None
    try:
        if math.isnan(float(value)):
            return None
    except Exception:
        pass
    return float(value)


def _make_formats(wb):
    return {
        'title': wb.add_format({'bold': True, 'font_size': 20, 'font_color': WHITE, 'bg_color': NAVY, 'align': 'left', 'valign': 'vcenter'}),
        'section': wb.add_format({'bold': True, 'font_size': 12, 'font_color': NAVY, 'bg_color': LIGHT_GOLD, 'bottom': 1, 'bottom_color': GOLD}),
        'label': wb.add_format({'bold': True, 'font_color': DARK, 'bg_color': LIGHT_GREY}),
        'text': wb.add_format({'font_color': DARK}),
        'header': wb.add_format({'bold': True, 'font_color': WHITE, 'bg_color': NAVY, 'border': 1, 'border_color': WHITE, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True}),
        'body': wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'valign': 'vcenter'}),
        'body_center': wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'align': 'center', 'valign': 'vcenter'}),
        'one_dec': wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'align': 'center', 'num_format': '0.0'}),
        'two_dec': wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'align': 'center', 'num_format': '0.00'}),
        'note': wb.add_format({'font_size': 9, 'font_color': MID_GREY, 'italic': True, 'text_wrap': True}),
        'counted': wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'bg_color': GREEN, 'align': 'center'}),
        'range_label': wb.add_format({'bold': True, 'font_size': 11, 'font_color': MID_GREY, 'bg_color': PALE_BLUE, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D8DDE3'}),
        'range_value': wb.add_format({'bold': True, 'font_size': 20, 'font_color': NAVY, 'bg_color': WHITE, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D8DDE3'}),
        'logo_band': wb.add_format({'bg_color': WHITE, 'bottom': 2, 'bottom_color': GOLD}),
    }


def _atar_range_text(result: dict) -> str:
    central = result.get('central_atar')
    if central is None:
        return '<50'
    low = max(0.0, float(central) - 5.0)
    high = min(99.95, float(central) + 5.0)
    return f'{low:.2f}–{high:.2f}'


def _write_summary_sheet(wb, ws, student_meta: dict, input_rows: list[dict], result: dict, logo_file: Path, has_logo: bool, fmt: dict):
    ws.hide_gridlines(2)
    ws.set_column('A:A', 30)
    ws.set_column('B:D', 21)
    ws.set_row(0, 44)
    ws.set_row(1, 27)
    ws.set_row(2, 20)

    ws.merge_range('A1:D2', '', fmt['logo_band'])
    if has_logo:
        ws.insert_image('A1', str(logo_file), {'x_scale': 0.20, 'y_scale': 0.20, 'x_offset': 8, 'y_offset': 3, 'object_position': 1})
    else:
        school_fmt = wb.add_format({'bold': True, 'font_size': 18, 'font_color': NAVY})
        ws.write('A1', 'REDDAM HOUSE NORTH SHORE', school_fmt)
    ws.merge_range('A3:D3', 'Five-Year NSW HSC → ATAR Estimator | Prediction Summary', fmt['title'])

    ws.merge_range('A5:D5', 'Student details', fmt['section'])
    details = [
        ('Student', student_meta.get('student_name', '') or '—'),
        ('Year group', student_meta.get('year_group', '') or '—'),
        ('Student ID', student_meta.get('student_id', '') or '—'),
        ('Prediction date', student_meta.get('prediction_date', '') or datetime.now().date().isoformat()),
    ]
    for i, (label, value) in enumerate(details, start=5):
        ws.write(i, 0, label, fmt['label'])
        ws.merge_range(i, 1, i, 3, value, fmt['text'])

    ws.merge_range('A11:D11', 'ATAR estimate', fmt['section'])
    ws.merge_range('A12:D12', 'Estimated ATAR range', fmt['range_label'])
    ws.merge_range('A13:D13', _atar_range_text(result), fmt['range_value'])
    ws.set_row(12, 36)

    ws.merge_range('A16:D16', 'Subject predictions', fmt['section'])
    headers = ['Course', 'Counted units', 'Mean scaled /50', 'ATAR contribution']
    for c, h in enumerate(headers):
        ws.write(16, c, h, fmt['header'])

    # Display subjects from highest to lowest ATAR contribution.
    # Subjects with no numeric ATAR contribution (for example <50) are placed last.
    subjects = sorted(
        result.get('subjects', []),
        key=lambda s: (
            s.get('atar_contribution') is not None,
            float(s.get('atar_contribution')) if s.get('atar_contribution') is not None else float('-inf'),
        ),
        reverse=True,
    )
    counts = result.get('counts', {})
    for r_idx, s in enumerate(subjects, start=17):
        course = s['course']
        counted_units = int(counts.get(course, 0))
        display_course = course + ('*' if s.get('provisional_proxy') else '')
        ws.write(r_idx, 0, display_course, fmt['body'])
        ws.write_number(r_idx, 1, counted_units, fmt['counted'] if counted_units > 0 else fmt['body_center'])
        ws.write_number(r_idx, 2, float(s.get('mean_scaled_per_unit')), fmt['two_dec'])
        contribution = s.get('atar_contribution')
        if contribution is None:
            ws.write(r_idx, 3, '<50', fmt['body_center'])
        else:
            ws.write_number(r_idx, 3, float(contribution), fmt['two_dec'])

    end_row = 17 + max(1, len(subjects))
    uses_hms_proxy = any(bool(s.get('provisional_proxy')) for s in subjects)
    summary_note = (
        'Estimate only. Teacher-predicted HSC bands (including Extension E-band codes) are converted to representative HSC-equivalent marks and then estimated using historical UAC scaling relationships. '
        'UAC notes that published HSC/scaled percentile data are not a one-to-one conversion for individual students.'
    )
    if uses_hms_proxy:
        summary_note += ' *Health and Movement Science is provisional and uses the 2021–2025 PDH&PE UAC scaling relationship as a proxy.'
    ws.merge_range(end_row + 2, 0, end_row + 3, 3, summary_note, fmt['note'])
    ws.freeze_panes(17, 0)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.set_margins(0.35, 0.35, 0.5, 0.5)


def _sheet_name_for_student(meta: dict, used: set[str]) -> str:
    surname = str(meta.get('surname', '') or '').strip()
    forename = str(meta.get('forename', '') or '').strip()
    if surname and forename:
        base = f'{surname}, {forename}'
    else:
        base = str(meta.get('student_name', '') or 'Student').strip()
    base = re.sub(r'[\[\]:*?/\\]', '', base).strip() or 'Student'
    base = base[:31]
    candidate = base
    n = 2
    while candidate.casefold() in used:
        suffix = f' ({n})'
        candidate = (base[:31 - len(suffix)] + suffix).strip()
        n += 1
    used.add(candidate.casefold())
    return candidate


def build_bulk_excel_export(reports: list[dict], logo_path: str | Path | None = None) -> bytes:
    """Create one workbook containing one Summary-style worksheet per student."""
    if not reports:
        raise ValueError('No student reports were supplied for the cohort export.')

    output = BytesIO()
    logo_file = Path(logo_path) if logo_path else Path(__file__).resolve().parent / 'assets' / 'reddam_house_north_shore_logo.png'
    has_logo = logo_file.exists()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    wb.set_properties({
        'title': 'Reddam House North Shore Cohort ATAR Estimates',
        'subject': 'Five-year NSW HSC to ATAR prediction summaries',
        'author': 'Reddam House North Shore',
        'company': 'Reddam House North Shore',
        'comments': 'Estimate based on published UAC scaling and aggregate data, 2021–2025.',
    })
    fmt = _make_formats(wb)
    used_names: set[str] = set()

    sorted_reports = sorted(
        reports,
        key=lambda r: (
            str(r.get('student_meta', {}).get('surname', '')).casefold(),
            str(r.get('student_meta', {}).get('forename', '')).casefold(),
            str(r.get('student_meta', {}).get('student_name', '')).casefold(),
        ),
    )
    for report in sorted_reports:
        meta = report['student_meta']
        sheet_name = _sheet_name_for_student(meta, used_names)
        ws = wb.add_worksheet(sheet_name)
        _write_summary_sheet(wb, ws, meta, report['input_rows'], report['result'], logo_file, has_logo, fmt)

    wb.close()
    output.seek(0)
    return output.getvalue()


def build_excel_export(student_meta: dict, input_rows: list[dict], result: dict, logo_path: str | Path | None = None) -> bytes:
    """Create the detailed individual-student .xlsx export."""
    output = BytesIO()
    logo_file = Path(logo_path) if logo_path else Path(__file__).resolve().parent / 'assets' / 'reddam_house_north_shore_logo.png'
    has_logo = logo_file.exists()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    wb.set_properties({
        'title': 'Reddam House North Shore ATAR Estimate',
        'subject': 'Five-year NSW HSC to ATAR estimate',
        'author': 'Reddam House North Shore',
        'company': 'Reddam House North Shore',
        'comments': 'Estimate based on published UAC scaling and aggregate data, 2021–2025.',
    })
    fmt = _make_formats(wb)

    ws = wb.add_worksheet('Summary')
    _write_summary_sheet(wb, ws, student_meta, input_rows, result, logo_file, has_logo, fmt)

    band_map = {str(r.get('course')): str(r.get('band', '') or '') for r in input_rows}
    subjects = result.get('subjects', [])

    # Scaling Detail
    ws2 = wb.add_worksheet('Scaling Detail')
    ws2.hide_gridlines(2)
    ws2.set_column('A:A', 27)
    ws2.set_column('B:B', 17)
    ws2.set_column('C:I', 15)
    ws2.merge_range('A1:I1', 'REDDAM HOUSE NORTH SHORE — Scaling Detail', fmt['title'])
    headers = ['Course', 'Predicted band', 'Converted HSC mark /100 equivalent'] + [str(y) for y in result.get('years', [])] + ['5-year mean']
    for c, h in enumerate(headers):
        ws2.write(2, c, h, fmt['header'])
    for r_idx, s in enumerate(subjects, start=3):
        display_course = s['course'] + ('*' if s.get('provisional_proxy') else '')
        ws2.write(r_idx, 0, display_course, fmt['body'])
        ws2.write(r_idx, 1, band_map.get(s['course'], ''), fmt['body_center'])
        ws2.write_number(r_idx, 2, float(s['predicted_hsc']), fmt['one_dec'])
        c = 3
        for y in result.get('years', []):
            v = _safe_num(s.get(f'scaled_{y}'))
            if v is None:
                ws2.write_blank(r_idx, c, None, fmt['body_center'])
            else:
                ws2.write_number(r_idx, c, v, fmt['two_dec'])
            c += 1
        ws2.write_number(r_idx, c, float(s['mean_scaled_per_unit']), fmt['two_dec'])
    detail_note = 'Values are estimated scaled marks per unit (/50).'
    if any(bool(s.get('provisional_proxy')) for s in subjects):
        detail_note += ' *HMS uses 2021–2025 PDH&PE scaling as a provisional proxy.'
    ws2.write(len(subjects) + 5, 0, detail_note, fmt['note'])
    ws2.freeze_panes(3, 3)

    # Historical scenarios
    ws3 = wb.add_worksheet('ATAR Scenarios')
    ws3.hide_gridlines(2)
    ws3.set_column('A:A', 14)
    ws3.set_column('B:C', 20)
    ws3.merge_range('A1:C1', 'REDDAM HOUSE NORTH SHORE — Historical ATAR Scenarios', fmt['title'])
    for c, h in enumerate(['Year', '10-unit aggregate', 'Estimated ATAR']):
        ws3.write(2, c, h, fmt['header'])
    for r_idx, row in enumerate(result.get('yearly', []), start=3):
        ws3.write_number(r_idx, 0, int(row['year']), fmt['body_center'])
        ws3.write_number(r_idx, 1, float(row['aggregate']), fmt['one_dec'])
        if row.get('atar') is None:
            ws3.write(r_idx, 2, '<50', fmt['body_center'])
        else:
            ws3.write_number(r_idx, 2, float(row['atar']), fmt['two_dec'])
    ws3.write(len(result.get('yearly', [])) + 5, 0, 'Each row applies that year’s available course-scaling relationships and aggregate→ATAR relationship.', fmt['note'])

    # Counted units
    ws4 = wb.add_worksheet('Counted Units')
    ws4.hide_gridlines(2)
    ws4.set_column('A:A', 27)
    ws4.set_column('B:B', 18)
    ws4.set_column('C:C', 20)
    ws4.set_column('D:D', 16)
    ws4.merge_range('A1:D1', 'REDDAM HOUSE NORTH SHORE — Units Counted in ATAR', fmt['title'])
    for c, h in enumerate(['Course', 'Unit within course', 'Scaled mark /50', 'English unit']):
        ws4.write(2, c, h, fmt['header'])
    for r_idx, row in enumerate(result.get('chosen_units', []), start=3):
        display_course = row['course'] + ('*' if row['course'] == HMS_COURSE else '')
        ws4.write(r_idx, 0, display_course, fmt['body'])
        ws4.write_number(r_idx, 1, int(row['unit_index']), fmt['body_center'])
        ws4.write_number(r_idx, 2, float(row['scaled']), fmt['two_dec'])
        ws4.write(r_idx, 3, 'Yes' if row['english'] else 'No', fmt['body_center'])

    # Methodology
    ws5 = wb.add_worksheet('Methodology')
    ws5.hide_gridlines(2)
    ws5.set_column('A:A', 26)
    ws5.set_column('B:B', 90)
    ws5.merge_range('A1:B1', 'REDDAM HOUSE NORTH SHORE — Methodology', fmt['title'])
    methodology = [
        ('Input', 'Teacher-selected predicted HSC performance bands. Standard courses use codes such as 6A, 6B and 6C; Extension courses use E4A, E4B and E4C through E1A, E1B and E1C. The app converts these to representative marks on a 0–100 equivalent scale.'),
        ('Scaling estimate', 'For each course and year 2021–2025, the app linearly interpolates between the HSC/scaled percentile pairs published in UAC Table A3.'),
        ('HMS provisional proxy', 'Health and Movement Science has no HMS rows in the 2021–2025 scaling dataset used by this model. When HMS is selected, the estimator uses the 2021–2025 PDH&PE scaling relationship as a temporary proxy while keeping the course labelled as Health and Movement Science.'),
        ('Central scaled mark', 'Mean of the available annual scaled-mark estimates for the selected course.'),
        ('Counting rule', 'Best 2 units of English plus best 8 remaining eligible units.'),
        ('ATAR estimate', 'The selected 10-unit scaled aggregate is converted to an ATAR using the mean 2021–2025 aggregate→ATAR relationship.'),
        ('Sensitivity', f"The app also tests the estimate using ±{result.get('sensitivity', 2):g} marks on the 0–100 equivalent prediction scale. For 1-unit Extension courses, ±2 on this scale is equivalent to ±1 mark out of 50."),
        ('ATAR contribution', 'Diagnostic ATAR corresponding to an aggregate in which all 10 counted units performed at that subject’s estimated scaled-mark level.'),
        ('Caution', 'This is an estimator. UAC states that published Table A3 HSC/scaled percentile pairs are not a one-to-one conversion for individual students.'),
        ('Data source', 'Universities Admissions Centre (UAC), Reports on the Scaling of the NSW HSC, 2021–2025.'),
    ]
    for r, (label, text) in enumerate(methodology, start=2):
        ws5.write(r, 0, label, fmt['label'])
        wrap_fmt = wb.add_format({'font_color': DARK, 'text_wrap': True, 'valign': 'top'})
        ws5.write(r, 1, text, wrap_fmt)
        ws5.set_row(r, 36)

    wb.close()
    output.seek(0)
    return output.getvalue()
