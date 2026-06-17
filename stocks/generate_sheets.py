# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl

INPUT_PATH = r'C:\Workspace\accounting\stocks\stocks.xlsx'
OUTPUT_PATH = r'C:\Workspace\accounting\stocks\stocks_output.xlsx'
SKIP_SHEET = '예시 > 현금성자산'
SPC_COL = '수익증권/spc명'


def find_company_name(ws):
    """1행 H(8)~K(11) 열에서 회사이름 찾기"""
    for col in range(8, 12):  # H=8, I=9, J=10, K=11
        v = ws.cell(1, col).value
        if v is not None and str(v).strip():
            return str(v).strip()
    return ''


def extract_detail_data(ws):
    """13행 헤더, 14행~ 데이터 추출"""
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(13, c).value
        if v is not None:
            h = str(v).replace(' ▼', '').strip()
            headers.append((c, h))

    if not headers:
        return pd.DataFrame()

    rows = []
    for r in range(14, ws.max_row + 1):
        row_data = {}
        has_data = False
        for col_idx, col_name in headers:
            v = ws.cell(r, col_idx).value
            if v is not None and str(v).strip() and str(v).strip() != '-':
                has_data = True
            row_data[col_name] = v
        if has_data:
            first_col_val = ws.cell(r, headers[0][0]).value
            if first_col_val is not None and str(first_col_val).strip():
                rows.append(row_data)

    return pd.DataFrame(rows, columns=[h for _, h in headers])


def main():
    wb = openpyxl.load_workbook(INPUT_PATH)
    sheet_dataframes = {}

    for name in wb.sheetnames:
        if name == SKIP_SHEET:
            continue

        ws = wb[name]
        company_name = find_company_name(ws)
        df = extract_detail_data(ws)

        if df.empty:
            print(f'[SKIP] {name}: 데이터 없음')
            continue

        df[SPC_COL] = company_name
        sheet_dataframes[name] = df
        print(f'[OK] {name}: {len(df)}행, 회사명={company_name}')

    all_dfs = []
    for name, df in sheet_dataframes.items():
        df_copy = df.copy()
        df_copy.insert(0, '시트명', name)
        all_dfs.append(df_copy)

    combined = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

    with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
        if not combined.empty:
            combined.to_excel(writer, sheet_name='전체', index=False)

        for name, df in sheet_dataframes.items():
            safe_name = name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

    print(f'\n저장 완료: {OUTPUT_PATH}')
    print(f'시트 목록: 전체 + {list(sheet_dataframes.keys())}')


if __name__ == '__main__':
    main()
