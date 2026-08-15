# Reddam House North Shore — Five-Year NSW ATAR Estimator (Version 2.1)

A Streamlit app for internal school planning using teacher-predicted HSC marks and published UAC scaling/aggregate data from 2021–2025.

## Version 2 features

- Official Reddam House North Shore logo plus navy/gold branded interface
- Official logo carried through to the Excel report
- Student name, year group, optional student ID and prediction date
- Subject-teacher attribution for every HSC prediction
- Five-year historical scaling estimate for each course
- Best 2 units of English + best 8 remaining units selection
- Central ATAR estimate, historical scaling range and teacher-prediction sensitivity band
- Graham-style ATAR contribution diagnostic for each subject
- One-click formatted Excel report export with five sheets:
  - Summary
  - Scaling Detail
  - ATAR Scenarios
  - Counted Units
  - Methodology

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Input convention

Enter teacher-predicted final HSC course marks on a 0–100 equivalent scale. For a 1-unit extension course, use the percentage-equivalent mark, e.g. 45/50 = 90.

## Methodology

1. For each course/year from 2021–2025, interpolate between the HSC/scaled percentile pairs published in UAC Table A3.
2. Use the mean of available annual scaled estimates as the central scaled mark.
3. Select the best 2 units of English plus the best 8 other units.
4. Convert the 10-unit scaled aggregate to an ATAR using the mean 2021–2025 aggregate→ATAR relationship.
5. Also calculate individual historical-year scenarios where the selected courses have complete data.
6. Calculate an ATAR contribution for each subject as a diagnostic ATAR-equivalent level.

## Important limitation

This is an estimator. UAC states that Table A3 is not a one-to-one conversion between an individual student's reported HSC mark and scaled mark. Teacher-predicted HSC marks also introduce prediction uncertainty.

## Data source

Universities Admissions Centre (UAC), Reports on the Scaling of the NSW HSC, 2021–2025.
