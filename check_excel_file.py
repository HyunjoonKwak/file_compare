import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 1) 비교에서 완전히 제외할 시트
EXCLUDED_SHEETS = {"Cover", "test description", "Results"}

# 2) 특수 비교를 적용할 시트(열 비교만 수행)
HISTORY_SHEET_NAME = "History"

# 3) 실제로 병합에 사용할 "Test ID"라는 이름 (B5에서 시작)
TEST_ID_COLUMN = "Test ID"

# 4) 색상 정의
BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")   # 추가(파란색)
YELLOW_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # 삭제(노란색)
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 변경(빨간색)

def load_excel_sheets(file_path):
    """
    엑셀 파일을 읽어서 {시트이름: DataFrame} 형태로 반환.
    - "Cover", "test description", "Results"는 제외
    - History 시트인지 아닌지 여부만 구분하여 DataFrame을 만든다.
    - Test ID가 B5 셀에 있다고 가정하고, 5번째 행을 헤더로 사용.
    """
    xls = pd.ExcelFile(file_path)
    sheet_dict = {}
    
    for sheet_name in xls.sheet_names:
        # 1) 비교 제외 시트는 건너뛴다
        if sheet_name in EXCLUDED_SHEETS:
            continue
        
        # 2) DataFrame 읽기 (헤더를 5행으로 지정)
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=4)
        
        # 3) Test ID 열이 B5라고 가정하고 열 이름 변경
        if TEST_ID_COLUMN not in df.columns:
            # B5의 헤더 값을 강제로 "Test ID"로 지정
            original_col = df.columns[1]  # B열이 0-based index로는 1번
            df.rename(columns={original_col: TEST_ID_COLUMN}, inplace=True)
        
        sheet_dict[sheet_name] = df
    
    return sheet_dict

def compare_history_columns(df1, df2, wb):
    """
    History 시트 전용 비교:
    - 첫 번째 파일(df1) 대비 두 번째 파일(df2)에서 '새로 추가된 열'만 찾는다.
    - 추가된 열만 시트에 기록, 파란색 하이라이트.
    - 기존 열, 값 변경 등은 무시.
    """
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    
    added_cols = cols2 - cols1  # df2에만 존재하는 열
    if not added_cols:
        # 추가된 열이 없다면 시트를 생성하지 않고 반환
        return
    
    # 새 워크시트 (예: "History_AddedCols")
    ws = wb.create_sheet(title="History_AddedCols")
    
    # 추가된 열만 추출
    added_df = df2[list(added_cols)]
    
    # 헤더 기록
    ws.append(list(added_cols))
    
    # 내용 기록
    for _, row_data in added_df.iterrows():
        row_list = [row_data[col] for col in added_cols]
        ws.append(row_list)
    
    # 파란색 하이라이트
    max_r = ws.max_row
    max_c = ws.max_column
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            ws.cell(row=r, column=c).fill = BLUE_FILL

def compare_sheets_by_test_id(sheet_name, df1, df2, wb):
    """
    History 탭이 아닌 시트 전용:
    - 이미 'B5'가 "Test ID"로 rename 된 상태임.
    - Test ID를 기준으로 행 비교(outer merge).
    - 추가(파란색), 삭제(노란색), 변경(빨간색).
    """
    if TEST_ID_COLUMN not in df1.columns or TEST_ID_COLUMN not in df2.columns:
        # 만약 rename이 안 되었거나 열 개수가 부족한 경우
        ws = wb.create_sheet(title=f"{sheet_name}_NO_KEY")
        ws.cell(row=1, column=1).value = f"'{TEST_ID_COLUMN}' 열이 없어 비교 불가"
        return
    
    merged_df = pd.merge(
        df1, df2,
        on=TEST_ID_COLUMN, how="outer", indicator=True,
        suffixes=("_old", "_new")
    )
    
    ws = wb.create_sheet(title=sheet_name)
    
    # df1, df2에서 Test ID 제외한 컬럼 파악
    df1_cols = list(df1.columns)
    df1_cols.remove(TEST_ID_COLUMN)
    df2_cols = list(df2.columns)
    df2_cols.remove(TEST_ID_COLUMN)
    
    # 기록할 컬럼 순서
    output_cols = [TEST_ID_COLUMN]
    for col in df1_cols:
        output_cols.append(col + "_old")
    for col in df2_cols:
        output_cols.append(col + "_new")
    output_cols.append("_merge")
    
    # 실제 존재하는 컬럼만 필터
    final_cols = [c for c in output_cols if c in merged_df.columns]
    
    # 헤더 기록
    ws.append(final_cols)
    # 데이터 기록
    for _, row_data in merged_df[final_cols].iterrows():
        ws.append(row_data.tolist())
    
    # 색상 처리
    merge_col_idx = final_cols.index("_merge") + 1
    old_map = {}
    new_map = {}
    
    for idx, col_name in enumerate(final_cols, start=1):
        if col_name.endswith("_old"):
            old_map[col_name[:-4]] = idx
        elif col_name.endswith("_new"):
            new_map[col_name[:-4]] = idx
    
    max_r = ws.max_row
    max_c = ws.max_column
    
    for r in range(2, max_r + 1):
        merge_val = ws.cell(row=r, column=merge_col_idx).value
        
        if merge_val == "left_only":
            # df1에만 존재 → 삭제
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL
        elif merge_val == "right_only":
            # df2에만 존재 → 추가
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL
        else:
            # both → 변경 체크
            for orig_col in old_map:
                old_idx = old_map[orig_col]
                new_idx = new_map.get(orig_col)
                if not new_idx:
                    continue
                val_old = ws.cell(row=r, column=old_idx).value
                val_new = ws.cell(row=r, column=new_idx).value
                if val_old != val_new:
                    ws.cell(row=r, column=old_idx).fill = RED_FILL
                    ws.cell(row=r, column=new_idx).fill = RED_FILL

def compare_files_logic(file1, file2, output_file):
    """
    실제 비교 로직:
      - History 탭 : df1 대비 df2에서 새로 추가된 열만 표시
      - 그 외 탭(제외 목록 제외) : Test ID(B5 기준)로 행 비교
      - 시트가 한쪽에만 있으면 전체 추가/삭제
    """
    df1_dict = load_excel_sheets(file1)
    df2_dict = load_excel_sheets(file2)
    
    wb = Workbook()
    wb.remove(wb.active)
    
    df1_sheets = set(df1_dict.keys())
    df2_sheets = set(df2_dict.keys())
    
    # 공통 시트
    common_sheets = df1_sheets & df2_sheets
    
    # 1) History 시트 처리
    if HISTORY_SHEET_NAME in common_sheets:
        df1_hist = df1_dict[HISTORY_SHEET_NAME]
        df2_hist = df2_dict[HISTORY_SHEET_NAME]
        
        compare_history_columns(df1_hist, df2_hist, wb)
        
        # 일반 Test ID 비교 대상에서는 제외
        common_sheets.remove(HISTORY_SHEET_NAME)
    
    # 2) 나머지 공통 시트 → Test ID 비교
    for sheet_name in sorted(common_sheets):
        compare_sheets_by_test_id(sheet_name, df1_dict[sheet_name], df2_dict[sheet_name], wb)
    
    # 3) df1에만 있는 시트 => 삭제
    df1_only = df1_sheets - df2_sheets
    for sheet_name in sorted(df1_only):
        df1_sheet = df1_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")
        for row_data in df1_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 노란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL
    
    # 4) df2에만 있는 시트 => 추가
    df2_only = df2_sheets - df1_sheets
    for sheet_name in sorted(df2_only):
        df2_sheet = df2_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        for row_data in df2_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 파란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL
    
    # 결과 저장
    wb.save(output_file)

# ───────────────────────── #
#      GUI 인터페이스        #
# ───────────────────────── #

def select_first_file():
    path = filedialog.askopenfilename(
        title="Select the first Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        first_file_var.set(path)

def select_second_file():
    path = filedialog.askopenfilename(
        title="Select the second Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        second_file_var.set(path)

def do_compare():
    f1 = first_file_var.get()
    f2 = second_file_var.get()
    
    if not f1 or not f2:
        messagebox.showerror("Error", "Both files must be selected!")
        return
    
    output_file = filed
