import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 1) 비교에서 완전히 제외할 시트
EXCLUDED_SHEETS = {"Cover", "test description", "Results"}

# 2) 특수 비교를 적용할 시트(열 비교)
HISTORY_SHEET_NAME = "History"

# 3) 행 비교 시 사용할 고유 키 컬럼
TEST_ID_COLUMN = "Test ID"

# 4) 색상 정의
BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")   # 추가(파란색)
YELLOW_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # 삭제(노란색)
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 변경(빨간색)

def load_excel_sheets(file_path):
    """
    주어진 엑셀 파일에서 시트들을 읽어와서
    { 시트이름: DataFrame } 형태로 반환.
    (Cover, test description, Results는 제외)
    """
    xls = pd.ExcelFile(file_path)
    sheet_dict = {}
    
    for sheet_name in xls.sheet_names:
        # 비교 제외 시트는 건너뛴다
        if sheet_name in EXCLUDED_SHEETS:
            continue
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        sheet_dict[sheet_name] = df
    
    return sheet_dict

def compare_history_columns(df1, df2, wb):
    """
    History 시트 전용 비교:
    - df1 대비 df2에서 새롭게 추가된 열(Column)만 찾아서 표시.
    - 기존에 있던 열은 스킵, 값 변경도 무시.
    - 추가된 열만 결과 시트에 출력(파란색 하이라이트).
    """
    # df1, df2 각각의 컬럼 집합
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    
    # 새로 추가된 컬럼들(두 번째 파일에만 있는 열)
    added_cols = cols2 - cols1
    
    if not added_cols:
        # 추가된 열이 없다면 따로 시트를 만들지 않고 반환
        return
    
    # 새 워크시트 생성 (예: "History_AddedCols")
    ws = wb.create_sheet(title="History_AddedCols")
    
    # 추가된 열만 뽑아서 DataFrame 구성
    added_df = df2[list(added_cols)]
    
    # 헤더 기록
    ws.append(list(added_cols))
    
    # 내용 기록
    for row_idx, row_data in added_df.iterrows():
        # row_data는 시리즈, 추가된 열 순서대로 값을 꺼내 리스트로 변환
        row_list = [row_data[col] for col in added_cols]
        ws.append(row_list)
    
    # 파란색 하이라이트 (전체)
    max_r = ws.max_row
    max_c = ws.max_column
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            ws.cell(row=r, column=c).fill = BLUE_FILL

def compare_sheets_by_test_id(sheet_name, df1, df2, wb):
    """
    History 이외의 탭 전용:
    - Test ID 컬럼 기준으로 df1, df2를 비교
    - 삭제(노란색), 추가(파란색), 변경(빨간색) 표시
    """
    # Test ID 컬럼이 없으면 비교 불가 → 그냥 시트 하나 만들고 메시지 남김
    if TEST_ID_COLUMN not in df1.columns or TEST_ID_COLUMN not in df2.columns:
        ws = wb.create_sheet(title=f"{sheet_name}_NO_KEY")
        ws.cell(row=1, column=1).value = f"'{TEST_ID_COLUMN}' 열이 없어 비교 불가"
        return
    
    # merge (outer)로 두 DF 병합
    merged_df = pd.merge(
        df1, df2,
        on=TEST_ID_COLUMN, how="outer", indicator=True,
        suffixes=("_old", "_new")
    )
    
    # 새 시트 생성
    ws = wb.create_sheet(title=sheet_name)
    
    # df1, df2에서 Test ID 제외한 컬럼 목록
    df1_cols = list(df1.columns)
    df1_cols.remove(TEST_ID_COLUMN)
    
    df2_cols = list(df2.columns)
    df2_cols.remove(TEST_ID_COLUMN)
    
    # 엑셀에 기록할 컬럼 순서: [Test ID, df1_old들, df2_new들, _merge]
    output_cols = [TEST_ID_COLUMN]
    
    for col in df1_cols:
        output_cols.append(col + "_old")
    for col in df2_cols:
        output_cols.append(col + "_new")
    
    output_cols.append("_merge")
    
    # 실제로 merged_df에 존재하는 컬럼만 필터
    final_cols = [c for c in output_cols if c in merged_df.columns]
    
    # 1) 헤더 기록
    ws.append(final_cols)
    
    # 2) 데이터 기록
    for _, row_data in merged_df[final_cols].iterrows():
        ws.append(row_data.tolist())
    
    # 색상 처리
    merge_col_idx = final_cols.index("_merge") + 1  # openpyxl 1-based index
    # old/new 컬럼 매핑
    old_map = {}
    new_map = {}
    
    for idx, col_name in enumerate(final_cols, start=1):
        if col_name.endswith("_old"):
            original = col_name[:-4]  # 예: "Name_old" -> "Name"
            old_map[original] = idx
        elif col_name.endswith("_new"):
            original = col_name[:-4]
            new_map[original] = idx
    
    max_r = ws.max_row
    max_c = ws.max_column
    
    for r in range(2, max_r + 1):  # 2행부터 데이터
        merge_val = ws.cell(row=r, column=merge_col_idx).value
        
        if merge_val == "left_only":
            # df1에만 존재 → 삭제(노란색)
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL
        
        elif merge_val == "right_only":
            # df2에만 존재 → 추가(파란색)
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL
        
        else:
            # both → 각 컬럼 old/new 비교
            for original_col, old_idx in old_map.items():
                new_idx = new_map.get(original_col)
                if not new_idx:
                    continue
                val_old = ws.cell(row=r, column=old_idx).value
                val_new = ws.cell(row=r, column=new_idx).value
                
                if val_old != val_new:
                    ws.cell(row=r, column=old_idx).fill = RED_FILL
                    ws.cell(row=r, column=new_idx).fill = RED_FILL

def compare_files_logic(file1, file2, output_file):
    """
    실제 비교 수행 로직:
      1) History 탭 → 열(Column) 추가 여부만 비교
      2) 나머지 탭 (Cover, test description, Results 제외) → Test ID 기반 행 비교
      3) 시트가 한쪽에만 있으면 전체 추가/삭제 처리
    """
    df1_dict = load_excel_sheets(file1)
    df2_dict = load_excel_sheets(file2)
    
    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 삭제
    
    df1_sheets = set(df1_dict.keys())
    df2_sheets = set(df2_dict.keys())
    
    # ──────────────────────────────────────
    # 1) History 시트 처리 (열 비교 전용)
    # ──────────────────────────────────────
    in_both = df1_sheets & df2_sheets
    
    if HISTORY_SHEET_NAME in in_both:
        # History 시트가 양쪽에 모두 있을 때만 "열 비교"를 시도
        df1_hist = df1_dict[HISTORY_SHEET_NAME]
        df2_hist = df2_dict[HISTORY_SHEET_NAME]
        
        compare_history_columns(df1_hist, df2_hist, wb)
        
        # 이미 처리했으므로 이후 일반 비교에선 제외
        in_both.remove(HISTORY_SHEET_NAME)
    # 만약 History 탭이 한쪽만 있다면 → 아래 "df1_only_sheets / df2_only_sheets" 처리
    
    # ──────────────────────────────────────────────
    # 2) 공통 시트 중 History 외 나머지 → Test ID 비교
    # ──────────────────────────────────────────────
    for sheet_name in sorted(in_both):
        # History는 위에서 뺐으므로 여기서는 일반 Test ID 비교
        compare_sheets_by_test_id(sheet_name, df1_dict[sheet_name], df2_dict[sheet_name], wb)
    
    # ───────────────────────────────
    # 3) df1 전용 시트, df2 전용 시트
    # ───────────────────────────────
    df1_only_sheets = df1_sheets - df2_sheets
    df2_only_sheets = df2_sheets - df1_sheets
    
    # df1에만 있는 시트 → 전부 삭제된 것으로 간주
    for sheet_name in sorted(df1_only_sheets):
        df1_sheet = df1_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")
        # 내용 복사
        for row_data in df1_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 노란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL
    
    # df2에만 있는 시트 → 전부 추가된 것으로 간주
    for sheet_name in sorted(df2_only_sheets):
        df2_sheet = df2_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        # 내용 복사
        for row_data in df2_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 파란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL
    
    # 최종 결과 저장
    wb.save(output_file)


# ───────────────────────── #
#      GUI 인터페이스        #
# ───────────────────────── #

def select_first_file():
    file_path = filedialog.askopenfilename(
        title="Select the first Excel file", 
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        first_file_var.set(file_path)

def select_second_file():
    file_path = filedialog.askopenfilename(
        title="Select the second Excel file", 
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        second_file_var.set(file_path)

def do_compare():
    f1 = first_file_var.get()
    f2 = second_file_var.get()
    
    if not f1 or not f2:
        messagebox.showerror("Error", "Both files must be selected!")
        return
    
    output_file = filedialog.asksaveasfilename(
        title="Save the comparison result",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_file:
        return
    
    try:
        compare_files_logic(f1, f2, output_file)
        messagebox.showinfo("Success", f"Comparison complete!\nSaved to: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel Comparison: History(Columns) + Others(Test ID)")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack()

    tk.Label(frame, text="First Excel File:", font=("Arial", 12)).grid(row=0, column=0, sticky="e", pady=5)
    first_file_var = tk.StringVar()
    tk.Entry(frame, textvariable=first_file_var, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
    tk.Button(frame, text="Browse", command=select_first_file).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(frame, text="Second Excel File:", font=("Arial", 12)).grid(row=1, column=0, sticky="e", pady=5)
    second_file_var = tk.StringVar()
    tk.Entry(frame, textvariable=second_file_var, width=50, state="readonly").grid(row=1, column=1, padx=5, pady=5)
    tk.Button(frame, text="Browse", command=select_second_file).grid(row=1, column=2, padx=5, pady=5)

    compare_btn = tk.Button(frame, text="Compare Files", command=do_compare, font=("Arial", 12))
    compare_btn.grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()
