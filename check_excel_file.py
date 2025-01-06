import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 비교 제외할 시트
EXCLUDED_SHEETS = {"Cover", "test description", "Results"}

# 특수 처리할 시트 (열 비교)
HISTORY_SHEET_NAME = "History"

# Test ID의 위치 (B5 셀 기준)
TEST_ID_COLUMN = "Test ID"

# 색상 정의
BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")   # 추가된 항목
YELLOW_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # 삭제된 항목
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 변경된 항목

def load_excel_sheets(file_path):
    """
    엑셀 파일을 읽어서 {시트이름: DataFrame} 형태로 반환.
    - EXCLUDED_SHEETS는 제외.
    - 모든 시트의 B5를 Test ID로 설정.
    """
    xls = pd.ExcelFile(file_path)
    sheet_dict = {}

    for sheet_name in xls.sheet_names:
        if sheet_name in EXCLUDED_SHEETS:
            continue
        
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=4)  # 5번째 행을 헤더로 설정
        if TEST_ID_COLUMN not in df.columns:
            # B5의 값을 Test ID로 설정
            original_col = df.columns[1]  # B 열은 0-based index에서 1번
            df.rename(columns={original_col: TEST_ID_COLUMN}, inplace=True)
        
        sheet_dict[sheet_name] = df
    
    return sheet_dict

def compare_history_columns(df1, df2, wb):
    """
    History 시트 전용: 첫 번째 파일(df1) 대비 두 번째 파일(df2)에서 새로 추가된 열(Column)만 비교.
    """
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)

    added_cols = cols2 - cols1  # df2에만 존재하는 열
    if not added_cols:
        return  # 추가된 열이 없으면 아무 작업도 하지 않음

    ws = wb.create_sheet(title="History_AddedCols")
    added_df = df2[list(added_cols)]

    ws.append(list(added_cols))  # 헤더 추가
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
    History 이외의 시트 전용: Test ID 열을 기준으로 행 비교.
    """
    if TEST_ID_COLUMN not in df1.columns or TEST_ID_COLUMN not in df2.columns:
        ws = wb.create_sheet(title=f"{sheet_name}_NO_KEY")
        ws.cell(row=1, column=1).value = f"'{TEST_ID_COLUMN}' 열이 없어 비교 불가"
        return

    merged_df = pd.merge(
        df1, df2, on=TEST_ID_COLUMN, how="outer", indicator=True, suffixes=("_old", "_new")
    )

    ws = wb.create_sheet(title=sheet_name)
    df1_cols = list(df1.columns)
    df1_cols.remove(TEST_ID_COLUMN)
    df2_cols = list(df2.columns)
    df2_cols.remove(TEST_ID_COLUMN)

    output_cols = [TEST_ID_COLUMN]
    for col in df1_cols:
        output_cols.append(col + "_old")
    for col in df2_cols:
        output_cols.append(col + "_new")
    output_cols.append("_merge")

    final_cols = [c for c in output_cols if c in merged_df.columns]
    ws.append(final_cols)

    for _, row_data in merged_df[final_cols].iterrows():
        ws.append(row_data.tolist())

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
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL
        elif merge_val == "right_only":
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL
        else:
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
    전체 비교 로직 수행.
    """
    df1_dict = load_excel_sheets(file1)
    df2_dict = load_excel_sheets(file2)

    wb = Workbook()
    wb.remove(wb.active)

    df1_sheets = set(df1_dict.keys())
    df2_sheets = set(df2_dict.keys())
    common_sheets = df1_sheets & df2_sheets

    if HISTORY_SHEET_NAME in common_sheets:
        df1_hist = df1_dict[HISTORY_SHEET_NAME]
        df2_hist = df2_dict[HISTORY_SHEET_NAME]
        compare_history_columns(df1_hist, df2_hist, wb)
        common_sheets.remove(HISTORY_SHEET_NAME)

    for sheet_name in sorted(common_sheets):
        compare_sheets_by_test_id(sheet_name, df1_dict[sheet_name], df2_dict[sheet_name], wb)

    df1_only = df1_sheets - df2_sheets
    for sheet_name in sorted(df1_only):
        df1_sheet = df1_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")
        for row_data in df1_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL

    df2_only = df2_sheets - df1_sheets
    for sheet_name in sorted(df2_only):
        df2_sheet = df2_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        for row_data in df2_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL

    wb.save(output_file)

def select_first_file():
    path = filedialog.askopenfilename(title="Select the first Excel file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        first_file_var.set(path)

def select_second_file():
    path = filedialog.askopenfilename(title="Select the second Excel file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        second_file_var.set(path)

def do_compare():
    f1 = first_file_var.get()
    f2 = second_file_var.get()
    if not f1 or not f2:
        messagebox.showerror("Error", "Both files must be selected!")
        return
    output_file = filedialog.asksaveasfilename(title="Save the comparison result", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file:
        return
    try:
        compare_files_logic(f1, f2, output_file)
        messagebox.showinfo("Success", f"Comparison complete!\nSaved to: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel Comparison Tool - History & Test ID")

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
