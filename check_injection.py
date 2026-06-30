# check_injection.py
# Auto-detect reference file DataAnalysisSheet_*.xlsx

import os
import pandas as pd
from collections import Counter
from openpyxl.styles import Font

REFERENCE_FILE = None


def find_reference_file():
    candidates = [
        f for f in os.listdir('.')
        if f.startswith('DataAnalysisSheet_')
        and f.endswith('.xlsx')
        and not f.startswith('~$')
    ]

    if len(candidates) == 0:
        raise FileNotFoundError(
            "No reference file matching 'DataAnalysisSheet_*.xlsx' found"
        )

    if len(candidates) > 1:
        raise RuntimeError(
            f"Multiple reference files found: {candidates}"
        )

    return candidates[0]


def get_report_filename():
    base = 'Injection_Check_Report'
    ext = '.xlsx'

    if not os.path.exists(base + ext):
        return base + ext

    i = 1
    while True:
        candidate = f'{base}_{i:02d}{ext}'
        if not os.path.exists(candidate):
            return candidate
        i += 1


def is_raw_file(filename):
    global REFERENCE_FILE

    name = filename.upper()

    if filename.startswith('~$'):
        return False

    if 'INJECTION_CHECK_REPORT' in name:
        return False

    if REFERENCE_FILE and filename == REFERENCE_FILE:
        return False

    return (
        'PROTEIN' in name
        or 'TG' in name
        or 'HB' in name
        or 'RF' in name
        or 'BR' in name
    )


def normalize_triplet(row):
    return (
        round(float(row[0]), 10),
        round(float(row[1]), 10),
        round(float(row[2]), 10)
    )


def detect_interferent(filename):
    name = filename.upper()

    if 'BR CONJ' in name or 'BRCONJ' in name:
        return 'Bilirubin conjugated'
    if 'PROTEIN' in name:
        return 'Protein'
    if 'TG_A2' in name or '_TG_' in name:
        return 'Triglycerides'
    if 'HB_A2' in name or '_HB_' in name:
        return 'Hemoglobin'
    if 'RF_A2' in name or '_RF_' in name:
        return 'Rheumatoid Factor'
    if 'BR_A2' in name or '_BR_' in name:
        return 'Bilirubin unconjugated'

    return 'UNKNOWN'


def extract_raw_triplets(file_path):
    df = pd.read_excel(file_path, header=None, engine='openpyxl')
    df = df.dropna(how='all')
    df = df.iloc[2:-2]
    df = df.iloc[:, [10, 11, 12]]
    df = df.dropna()

    triplets = [normalize_triplet(r) for r in df.values]
    return Counter(triplets), triplets


def extract_reference_triplets(reference_df, interferent):
    df = reference_df.copy()
    df['Interferent'] = df['Interferent'].ffill()

    df = df[df['Interferent'].astype(str).str.strip().eq(interferent)]
    df = df[['Sample OD', 'OD c/o', 'Value']]
    df = df.dropna(subset=['Sample OD', 'OD c/o', 'Value'])

    triplets = [normalize_triplet(r) for r in df.values]
    return Counter(triplets), triplets


def autosize(ws):
    for column in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in column)
        ws.column_dimensions[column[0].column_letter].width = min(width + 4, 80)


def main():
    global REFERENCE_FILE

    REFERENCE_FILE = find_reference_file()
    print(f'Reference file detected: {REFERENCE_FILE}')

    reference_df = pd.read_excel(
        REFERENCE_FILE,
        sheet_name='Data Entry',
        engine='openpyxl'
    )

    files = [
        f for f in os.listdir('.')
        if f.endswith('.xlsx') and is_raw_file(f)
    ]

    summary = []
    details = []
    missing_rows = []
    extra_rows = []

    for file in sorted(files):
        interferent = detect_interferent(file)

        raw_counter, raw_triplets = extract_raw_triplets(file)
        ref_counter, ref_triplets = extract_reference_triplets(reference_df, interferent)

        missing = raw_counter - ref_counter
        extra = ref_counter - raw_counter

        missing_count = sum(missing.values())
        extra_count = sum(extra.values())

        status = '✅ OK' if missing_count == 0 and extra_count == 0 else '⚠ CHECK REQUIRED'

        summary.append([
            file, interferent, status,
            len(raw_triplets), len(ref_triplets),
            missing_count, extra_count
        ])

        details.append([
            file, interferent, len(raw_triplets), len(ref_triplets),
            missing_count, extra_count, status
        ])

        for t, c in missing.items():
            missing_rows.append([file, interferent, t[0], t[1], t[2], c])

        for t, c in extra.items():
            extra_rows.append([file, interferent, t[0], t[1], t[2], c])

    report_file = get_report_filename()

    with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
        pd.DataFrame(summary, columns=[
            'File','Interferent','Status','Raw Triplets',
            'Reference Triplets','Missing','Extra'
        ]).to_excel(writer, sheet_name='Summary', index=False)

        pd.DataFrame(details, columns=[
            'File','Interferent','Raw Count','Reference Count',
            'Missing','Extra','Status'
        ]).to_excel(writer, sheet_name='Detailed Comparison', index=False)

        pd.DataFrame(missing_rows, columns=[
            'File','Interferent','Sample OD','OD c/o','Value','Count'
        ]).to_excel(writer, sheet_name='Missing Triplets', index=False)

        pd.DataFrame(extra_rows, columns=[
            'File','Interferent','Sample OD','OD c/o','Value','Count'
        ]).to_excel(writer, sheet_name='Extra Triplets', index=False)

        for ws in writer.book.worksheets:
            autosize(ws)
            for cell in ws[1]:
                cell.font = Font(bold=True)

    print(f'Report generated: {report_file}')


if __name__ == '__main__':
    main()
