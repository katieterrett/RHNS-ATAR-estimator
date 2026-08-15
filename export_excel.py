from __future__ import annotations

from io import BytesIO
from datetime import datetime
from pathlib import Path
import math
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


def _safe_num(value):
    if value is None:
        return None
    try:
        if math.isnan(float(value)):
            return None
    except Exception:
        pass
    return float(value)


def build_excel_export(student_meta: dict, input_rows: list[dict], result: dict, logo_path: str | Path | None = None) -> bytes:
    """Create a polished .xlsx export for the current student estimate."""
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

    # Common formats
    title_fmt = wb.add_format({'bold': True, 'font_size': 20, 'font_color': WHITE, 'bg_color': NAVY, 'align': 'left', 'valign': 'vcenter'})
    subtitle_fmt = wb.add_format({'font_size': 10, 'font_color': '#D7DDE5', 'bg_color': NAVY, 'align': 'left', 'valign': 'vcenter'})
    section_fmt = wb.add_format({'bold': True, 'font_size': 12, 'font_color': NAVY, 'bg_color': LIGHT_GOLD, 'bottom': 1, 'bottom_color': GOLD})
    label_fmt = wb.add_format({'bold': True, 'font_color': DARK, 'bg_color': LIGHT_GREY})
    text_fmt = wb.add_format({'font_color': DARK})
    header_fmt = wb.add_format({'bold': True, 'font_color': WHITE, 'bg_color': NAVY, 'border': 1, 'border_color': WHITE, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    body_fmt = wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'valign': 'vcenter'})
    body_center_fmt = wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'align': 'center', 'valign': 'vcenter'})
    one_dec_fmt = wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'align': 'center', 'num_format': '0.0'})
    two_dec_fmt = wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'align': 'center', 'num_format': '0.00'})
    metric_label_fmt = wb.add_format({'bold': True, 'font_color': MID_GREY, 'bg_color': PALE_BLUE, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D8DDE3'})
    metric_value_fmt = wb.add_format({'bold': True, 'font_size': 18, 'font_color': NAVY, 'bg_color': WHITE, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D8DDE3', 'num_format': '0.00'})
    note_fmt = wb.add_format({'font_size': 9, 'font_color': MID_GREY, 'italic': True, 'text_wrap': True})
    counted_fmt = wb.add_format({'border': 1, 'border_color': '#D8DDE3', 'bg_color': GREEN, 'align': 'center'})

    # ---------------- Summary ----------------
    ws = wb.add_worksheet('Summary')
    ws.hide_gridlines(2)
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 26)
    ws.set_column('C:H', 17)
    ws.set_row(0, 44)
    ws.set_row(1, 27)
    ws.set_row(2, 20)
    # White logo band followed by the navy report title.
    ws.merge_range('A1:H2', '', wb.add_format({'bg_color': WHITE, 'bottom': 2, 'bottom_color': GOLD}))
    if has_logo:
        ws.insert_image('A1', str(logo_file), {'x_scale': 0.20, 'y_scale': 0.20, 'x_offset': 8, 'y_offset': 3, 'object_position': 1})
    else:
        ws.write('A1', 'REDDAM HOUSE NORTH SHORE', wb.add_format({'bold': True, 'font_size': 18, 'font_color': NAVY}))
    ws.merge_range('A3:H3', 'Five-Year NSW HSC → ATAR Estimator | Teacher Prediction Report', title_fmt)

    ws.merge_range('A5:H5', 'Student details', section_fmt)
    details = [
        ('Student', student_meta.get('student_name', '') or '—'),
        ('Year group', student_meta.get('year_group', '') or '—'),
        ('Student ID', student_meta.get('student_id', '') or '—'),
        ('Prediction date', student_meta.get('prediction_date', '') or datetime.now().date().isoformat()),
    ]
    for i, (label, value) in enumerate(details, start=5):
        ws.write(i, 0, label, label_fmt)
        ws.merge_range(i, 1, i, 3, value, text_fmt)

    central = result.get('central_atar')
    yearly_atars = [x['atar'] for x in result.get('yearly', []) if x.get('atar') is not None]
    hist_low = min(yearly_atars) if yearly_atars else None
    hist_high = max(yearly_atars) if yearly_atars else None
    sens_low = result.get('low', {}).get('atar')
    sens_high = result.get('high', {}).get('atar')

    ws.merge_range('A11:H11', 'ATAR estimate', section_fmt)
    metrics = [
        ('Central ATAR', central),
        ('10-unit aggregate', result.get('central_aggregate')),
        ('Historical low', hist_low),
        ('Historical high', hist_high),
    ]
    for col, (label, value) in enumerate(metrics):
        c1 = col * 2
        ws.merge_range(11, c1, 11, c1 + 1, label, metric_label_fmt)
        if value is None:
            ws.merge_range(12, c1, 12, c1 + 1, '<50', metric_value_fmt)
        else:
            ws.merge_range(12, c1, 12, c1 + 1, value, metric_value_fmt)
    ws.set_row(12, 32)

    ws.merge_range('A15:H15', 'Prediction sensitivity', section_fmt)
    ws.write('A16', 'Sensitivity assumption', label_fmt)
    ws.write('B16', f"±{result.get('sensitivity', 0):g} HSC marks", text_fmt)
    ws.write('D16', 'Low scenario ATAR', label_fmt)
    ws.write('E16', '<50' if sens_low is None else sens_low, text_fmt)
    ws.write('F16', 'High scenario ATAR', label_fmt)
    ws.write('G16', '<50' if sens_high is None else sens_high, text_fmt)

    ws.merge_range('A19:H19', 'Subject predictions', section_fmt)
    summary_headers = ['Course', 'Subject teacher', 'Predicted HSC', 'Units', 'Counted units', 'Mean scaled /50', 'Scaled range /50', 'ATAR contribution']
    for c, h in enumerate(summary_headers):
        ws.write(19, c, h, header_fmt)

    teacher_map = {str(r.get('course')): str(r.get('teacher', '') or '') for r in input_rows}
    subjects = result.get('subjects', [])
    counts = result.get('counts', {})
    for r_idx, s in enumerate(subjects, start=20):
        course = s['course']
        vals = [
            course,
            teacher_map.get(course, ''),
            s.get('predicted_hsc'),
            s.get('units'),
            counts.get(course, 0),
            s.get('mean_scaled_per_unit'),
            f"{s.get('scaled_min', 0):.2f}–{s.get('scaled_max', 0):.2f}",
            s.get('atar_contribution'),
        ]
        ws.write(r_idx, 0, vals[0], body_fmt)
        ws.write(r_idx, 1, vals[1], body_fmt)
        ws.write_number(r_idx, 2, float(vals[2]), one_dec_fmt)
        ws.write_number(r_idx, 3, int(vals[3]), body_center_fmt)
        if int(vals[4]) > 0:
            ws.write_number(r_idx, 4, int(vals[4]), counted_fmt)
        else:
            ws.write_number(r_idx, 4, int(vals[4]), body_center_fmt)
        ws.write_number(r_idx, 5, float(vals[5]), two_dec_fmt)
        ws.write(r_idx, 6, vals[6], body_center_fmt)
        if vals[7] is None:
            ws.write(r_idx, 7, '<50', body_center_fmt)
        else:
            ws.write_number(r_idx, 7, float(vals[7]), two_dec_fmt)

    end_row = 20 + max(1, len(subjects))
    ws.merge_range(end_row + 2, 0, end_row + 3, 7,
                   'Estimate only. Teacher-predicted HSC marks are converted using historical UAC scaling relationships. '
                   'UAC notes that published HSC/scaled percentile data are not a one-to-one conversion for individual students.',
                   note_fmt)
    ws.freeze_panes(20, 0)

    # ---------------- Annual scaling ----------------
    ws2 = wb.add_worksheet('Scaling Detail')
    ws2.hide_gridlines(2)
    ws2.set_column('A:A', 27)
    ws2.set_column('B:B', 23)
    ws2.set_column('C:I', 15)
    ws2.merge_range('A1:I1', 'REDDAM HOUSE NORTH SHORE — Scaling Detail', title_fmt)
    if has_logo:
        ws2.set_row(0, 30)
        ws2.insert_image('H1', str(logo_file), {'x_scale': 0.10, 'y_scale': 0.10, 'x_offset': 5, 'y_offset': 2, 'object_position': 1})
    headers = ['Course', 'Subject teacher', 'Predicted HSC'] + [str(y) for y in result.get('years', [])] + ['5-year mean']
    for c, h in enumerate(headers):
        ws2.write(2, c, h, header_fmt)
    for r_idx, s in enumerate(subjects, start=3):
        ws2.write(r_idx, 0, s['course'], body_fmt)
        ws2.write(r_idx, 1, teacher_map.get(s['course'], ''), body_fmt)
        ws2.write_number(r_idx, 2, float(s['predicted_hsc']), one_dec_fmt)
        c = 3
        for y in result.get('years', []):
            v = _safe_num(s.get(f'scaled_{y}'))
            if v is None:
                ws2.write_blank(r_idx, c, None, body_center_fmt)
            else:
                ws2.write_number(r_idx, c, v, two_dec_fmt)
            c += 1
        ws2.write_number(r_idx, c, float(s['mean_scaled_per_unit']), two_dec_fmt)
    ws2.write(len(subjects) + 5, 0, 'Values are estimated scaled marks per unit (/50).', note_fmt)
    ws2.freeze_panes(3, 3)

    # ---------------- Historical scenarios ----------------
    ws3 = wb.add_worksheet('ATAR Scenarios')
    ws3.hide_gridlines(2)
    ws3.set_column('A:A', 14)
    ws3.set_column('B:C', 20)
    ws3.merge_range('A1:C1', 'REDDAM HOUSE NORTH SHORE — Historical ATAR Scenarios', title_fmt)
    if has_logo:
        ws3.set_row(0, 30)
        ws3.insert_image('C1', str(logo_file), {'x_scale': 0.08, 'y_scale': 0.08, 'x_offset': 3, 'y_offset': 2, 'object_position': 1})
    for c, h in enumerate(['Year', '10-unit aggregate', 'Estimated ATAR']):
        ws3.write(2, c, h, header_fmt)
    for r_idx, row in enumerate(result.get('yearly', []), start=3):
        ws3.write_number(r_idx, 0, int(row['year']), body_center_fmt)
        ws3.write_number(r_idx, 1, float(row['aggregate']), one_dec_fmt)
        if row.get('atar') is None:
            ws3.write(r_idx, 2, '<50', body_center_fmt)
        else:
            ws3.write_number(r_idx, 2, float(row['atar']), two_dec_fmt)
    ws3.write(len(result.get('yearly', [])) + 5, 0,
              'Each row applies that year’s available course-scaling relationships and aggregate→ATAR relationship.', note_fmt)

    # ---------------- Counted units ----------------
    ws4 = wb.add_worksheet('Counted Units')
    ws4.hide_gridlines(2)
    ws4.set_column('A:A', 27)
    ws4.set_column('B:B', 18)
    ws4.set_column('C:C', 20)
    ws4.set_column('D:D', 16)
    ws4.merge_range('A1:D1', 'REDDAM HOUSE NORTH SHORE — Units Counted in ATAR', title_fmt)
    if has_logo:
        ws4.set_row(0, 30)
        ws4.insert_image('D1', str(logo_file), {'x_scale': 0.08, 'y_scale': 0.08, 'x_offset': 3, 'y_offset': 2, 'object_position': 1})
    for c, h in enumerate(['Course', 'Unit within course', 'Scaled mark /50', 'English unit']):
        ws4.write(2, c, h, header_fmt)
    for r_idx, row in enumerate(result.get('chosen_units', []), start=3):
        ws4.write(r_idx, 0, row['course'], body_fmt)
        ws4.write_number(r_idx, 1, int(row['unit_index']), body_center_fmt)
        ws4.write_number(r_idx, 2, float(row['scaled']), two_dec_fmt)
        ws4.write(r_idx, 3, 'Yes' if row['english'] else 'No', body_center_fmt)

    # ---------------- Methodology ----------------
    ws5 = wb.add_worksheet('Methodology')
    ws5.hide_gridlines(2)
    ws5.set_column('A:A', 26)
    ws5.set_column('B:B', 90)
    ws5.merge_range('A1:B1', 'REDDAM HOUSE NORTH SHORE — Methodology', title_fmt)
    if has_logo:
        ws5.set_row(0, 30)
        ws5.insert_image('B1', str(logo_file), {'x_scale': 0.08, 'y_scale': 0.08, 'x_offset': 3, 'y_offset': 2, 'object_position': 1})
    methodology = [
        ('Input', 'Teacher-predicted final HSC course marks, entered on a 0–100 equivalent scale. For 1-unit extension courses, use the percentage-equivalent mark (e.g. 45/50 = 90).'),
        ('Scaling estimate', 'For each course and year 2021–2025, the app linearly interpolates between the HSC/scaled percentile pairs published in UAC Table A3.'),
        ('Central scaled mark', 'Mean of the available annual scaled-mark estimates for the selected course.'),
        ('Counting rule', 'Best 2 units of English plus best 8 remaining eligible units.'),
        ('ATAR estimate', 'The selected 10-unit scaled aggregate is converted to an ATAR using the mean 2021–2025 aggregate→ATAR relationship.'),
        ('Historical range', 'The app also recalculates the estimate using each historical year where all selected courses have usable data.'),
        ('ATAR contribution', 'Diagnostic ATAR corresponding to an aggregate in which all 10 counted units performed at that subject’s estimated scaled-mark level.'),
        ('Caution', 'This is an estimator. UAC states that published Table A3 HSC/scaled percentile pairs are not a one-to-one conversion for individual students.'),
        ('Data source', 'Universities Admissions Centre (UAC), Reports on the Scaling of the NSW HSC, 2021–2025.'),
    ]
    for r, (label, text) in enumerate(methodology, start=2):
        ws5.write(r, 0, label, label_fmt)
        ws5.write(r, 1, text, wb.add_format({'font_color': DARK, 'text_wrap': True, 'valign': 'top'}))
        ws5.set_row(r, 36)

    wb.close()
    output.seek(0)
    return output.getvalue()
