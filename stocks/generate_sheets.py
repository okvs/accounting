# -*- coding: utf-8 -*-
import glob
import os
import re

import pandas as pd
import openpyxl

STOCKS_DIR = r'C:\Workspace\accounting\stocks'
OUTPUT_PATH = os.path.join(STOCKS_DIR, 'stocks_output.xlsx')
SKIP_PATTERNS = ['예시']
SPC_COL = '수익증권/spc명'


def normalize_sheet_name(name):
    return re.sub(r'\s+', '', name)


def is_skip_sheet(name):
    for pat in SKIP_PATTERNS:
        if pat in name:
            return True
    return False


def find_company_name(ws):
    for col in range(8, 12):  # H=8 ~ K=11
        v = ws.cell(1, col).value
        if v is not None and str(v).strip():
            return str(v).strip()
    return ''


def find_header_row(ws):
    last_match = None
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column + 1, 10)):
            v = ws.cell(r, c).value
            if v and re.match(r'계정과목', str(v)):
                all_vals = [ws.cell(r, col).value for col in range(1, ws.max_column + 1)]
                non_none = [x for x in all_vals if x is not None]
                if len(non_none) >= 3:
                    last_match = r
    return last_match


def extract_detail_data(ws, header_row):
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).replace(' ▼', '').strip()
            headers.append((c, h))

    if not headers:
        return pd.DataFrame()

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        row_data = {}
        has_data = False
        for col_idx, col_name in headers:
            v = ws.cell(r, col_idx).value
            if v is not None and str(v).strip() and str(v).strip() != '-':
                has_data = True
            row_data[col_name] = v
        if has_data:
            all_text = ' '.join(str(v) for v in row_data.values() if v is not None)
            if re.search(r'\d+\s*페이지', all_text):
                continue
            rows.append(row_data)

    return pd.DataFrame(rows, columns=[h for _, h in headers])


def get_category_from_a1(ws):
    v = ws.cell(1, 1).value
    if v and '결산명세_' in str(v):
        return str(v).split('결산명세_', 1)[1].strip()
    return None


def main():
    xlsx_files = sorted(glob.glob(os.path.join(STOCKS_DIR, '*.xlsx')))
    xlsx_files = [f for f in xlsx_files if os.path.basename(f) != os.path.basename(OUTPUT_PATH)]

    if not xlsx_files:
        print('처리할 엑셀 파일이 없습니다.')
        return

    print(f'대상 파일 {len(xlsx_files)}개:')
    for f in xlsx_files:
        print(f'  - {os.path.basename(f)}')

    combined_by_category = {}

    for fpath in xlsx_files:
        fname = os.path.basename(fpath)
        print(f'\n=== {fname} ===')
        wb = openpyxl.load_workbook(fpath, data_only=True)

        for sheet_name in wb.sheetnames:
            if is_skip_sheet(sheet_name):
                print(f'  [SKIP] {sheet_name}: 예시 시트')
                continue

            ws = wb[sheet_name]

            category = get_category_from_a1(ws)
            if not category:
                category = normalize_sheet_name(sheet_name)

            company_name = find_company_name(ws)
            header_row = find_header_row(ws)

            if header_row is None:
                print(f'  [SKIP] {sheet_name}: 계정과목 헤더 없음')
                continue

            df = extract_detail_data(ws, header_row)

            if df.empty:
                print(f'  [SKIP] {sheet_name}: 데이터 없음')
                continue

            df[SPC_COL] = company_name
            norm_cat = normalize_sheet_name(category)

            if norm_cat not in combined_by_category:
                combined_by_category[norm_cat] = []
            combined_by_category[norm_cat].append(df)
            print(f'  [OK] {sheet_name} -> {norm_cat}: {len(df)}행, 회사명={company_name}')

        wb.close()

    if not combined_by_category:
        print('\n추출된 데이터가 없습니다.')
        return

    merged_sheets = {}
    for cat, dfs in combined_by_category.items():
        merged_sheets[cat] = pd.concat(dfs, ignore_index=True)

    all_dfs = []
    for cat, df in merged_sheets.items():
        df_copy = df.copy()
        df_copy.insert(0, '시트명', cat)
        all_dfs.append(df_copy)
    combined_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

    with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
        if not combined_all.empty:
            combined_all.to_excel(writer, sheet_name='전체', index=False)

        for cat, df in merged_sheets.items():
            safe_name = cat[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

    print(f'\n저장 완료: {OUTPUT_PATH}')
    print(f'시트 목록: 전체 + {list(merged_sheets.keys())}')


if __name__ == '__main__':
    main()
