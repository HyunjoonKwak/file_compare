import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 비교 제외할 시트 (이름이 정확히 일치하는 경우 제외)
EXCLUDED_SHEETS = {"Cover", "test description", "Results"}

# 색상 정의
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 변경된 셀
BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")   # 추가된 행
YELLOW_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # 삭제된 행

# "Test ID"를 기본 키 컬럼으로 사용
TEST_ID_COLUMN = "Test ID"

def compare_two_sheets_with_key(sheet_name, df1_sheet, df2_sheet, wb, key_column=TEST_ID_COLUMN):
    """
    key_column(기본값: 'Test ID')을 기준으로 df1_sheet와 df2_sheet를 비교.
    결과를 wb(Workbook)에 새 시트(sheet_name)로 기록.
    """
    # outer join으로 병합하여, 어떤 행이 왼쪽(df1), 오른쪽(df2)에만 있는지 표시(_merge)
    merged_df = pd.merge(
        df1_sheet, df2_sheet,
        on=key_column, how="outer", indicator=True,
        suffixes=("_old", "_new")
    )

    # 시트 생성
    ws = wb.create_sheet(title=sheet_name)

    # df1과 df2의 컬럼명을 각각 파악
    df1_cols = list(df1_sheet.columns)
    df2_cols = list(df2_sheet.columns)

    # key_column은 맨 앞으로 빼고, 나머지 컬럼을 old/new 순으로 배치
    df1_cols.remove(key_column)
    df2_cols.remove(key_column)

    # 최종으로 엑셀에 기록할 컬럼 순서
    # [키컬럼, df1_old, df2_new, ..., _merge]
    output_cols = [key_column]
    for col in df1_cols:
        output_cols.append(col + "_old")
    for col in df2_cols:
        output_cols.append(col + "_new")
    output_cols.append("_merge")

    # 실제 merged_df에 존재하는 컬럼(병합 과정에서 누락될 수 있는 컬럼 제외)
    final_cols = [c for c in output_cols if c in merged_df.columns]

    # 1) 헤더 기록
    ws.append(final_cols)

    # 2) 데이터 기록
    for _, row_data in merged_df[final_cols].iterrows():
        ws.append(row_data.tolist())

    # _merge 컬럼의 위치(1-based index)
    merge_col_idx = final_cols.index("_merge") + 1

    # old/new 컬럼 위치를 찾기 위한 매핑
    old_map = {}  # 예: {"Name": 시트에서 Name_old 컬럼 인덱스}
    new_map = {}  # 예: {"Name": 시트에서 Name_new 컬럼 인덱스}

    for idx, col_name in enumerate(final_cols, start=1):
        if col_name.endswith("_old"):
            # 예: "Name_old" → "Name"
            original_col = col_name[:-4]
            old_map[original_col] = idx
        elif col_name.endswith("_new"):
            # 예: "Name_new" → "Name"
            original_col = col_name[:-4]
            new_map[original_col] = idx

    # 3) 색상 하이라이트
    max_row = ws.max_row
    max_col = ws.max_column

    for row_idx in range(2, max_row + 1):  # 2행부터 실제 데이터
        merge_value = ws.cell(row=row_idx, column=merge_col_idx).value

        if merge_value == "left_only":
            # df1에만 존재 -> 삭제된 행
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = YELLOW_FILL

        elif merge_value == "right_only":
            # df2에만 존재 -> 추가된 행
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = BLUE_FILL

        else:
            # both -> 셀 단위 비교 (old != new)
            for orig_col, old_col_idx in old_map.items():
                new_col_idx = new_map.get(orig_col)
                if not new_col_idx:
                    continue  # df2에 없는 컬럼이면 비교 불가

                val_old = ws.cell(row=row_idx, column=old_col_idx).value
                val_new = ws.cell(row=row_idx, column=new_col_idx).value

                # 값이 다르면 변경(빨간색)
                if val_old != val_new:
                    ws.cell(row=row_idx, column=old_col_idx).fill = RED_FILL
                    ws.cell(row=row_idx, column=new_col_idx).fill = RED_FILL

def highlight_changes_with_test_id(file1, file2, output_file):
    """
    두 엑셀 파일(file1, file2)을 읽어,
    - EXCLUDED_SHEETS( Cover, test description, Results )는 건너뛰고
    - 나머지 시트 중 공통된 시트는 'Test ID' 컬럼 기준으로 행 매칭 비교
    - df1에만 있는 시트 -> 전체 삭제(노란색)로 표시
    - df2에만 있는 시트 -> 전체 추가(파란색)로 표시
    결과를 output_file에 저장.
    """
    df1 = pd.read_excel(file1, sheet_name=None)
    df2 = pd.read_excel(file2, sheet_name=None)

    wb = Workbook()
    # 기본 생성 시트 제거
    wb.remove(wb.active)

    df1_sheets = set(df1.keys())
    df2_sheets = set(df2.keys())

    # 비교 대상 시트: 공통 시트 중에서 제외 목록에 없는 시트
    common_sheets = (df1_sheets & df2_sheets) - EXCLUDED_SHEETS

    for sheet_name in common_sheets:
        # 시트에 "Test ID" 컬럼이 없으면 스킵하거나 다른 처리
        if TEST_ID_COLUMN not in df1[sheet_name].columns or TEST_ID_COLUMN not in df2[sheet_name].columns:
            ws = wb.create_sheet(title=f"{sheet_name}_NO_KEY")
            ws.cell(row=1, column=1, value="해당 시트에는 'Test ID' 컬럼이 없어 비교 불가")
            continue

        # 공통 시트 비교
        df1_sheet = df1[sheet_name]
        df2_sheet = df2[sheet_name]
        compare_two_sheets_with_key(sheet_name, df1_sheet, df2_sheet, wb, key_column=TEST_ID_COLUMN)

    # df1에만 있고, df2에는 없는 시트 -> 삭제(노란색)
    df1_only_sheets = (df1_sheets - df2_sheets) - EXCLUDED_SHEETS
    for sheet_name in df1_only_sheets:
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")
        df1_sheet = df1[sheet_name]
        for row_data in df1_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 모든 셀 노란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL

    # df2에만 있고, df1에는 없는 시트 -> 추가(파란색)
    df2_only_sheets = (df2_sheets - df1_sheets) - EXCLUDED_SHEETS
    for sheet_name in df2_only_sheets:
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        df2_sheet = df2[sheet_name]
        for row_data in df2_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 모든 셀 파란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL

    wb.save(output_file)

# ───────────────────────── #
#         GUI 인터페이스      #
# ───────────────────────── #

def select_first_file():
    file = filedialog.askopenfilename(
        title="Select the first Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file:
        first_file_var.set(file)

def select_second_file():
    file = filedialog.askopenfilename(
        title="Select the second Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file:
        second_file_var.set(file)

def compare_files():
    file1 = first_file_var.get()
    file2 = second_file_var.get()

    if not file1 or not file2:
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
        # "Test ID" 열을 기준으로 비교 수행
        highlight_changes_with_test_id(file1, file2, output_file)
        messagebox.showinfo("Success", f"Comparison complete!\nResult saved to: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel File Comparison (Key: 'Test ID')")

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

    compare_button = tk.Button(frame, text="Compare Files", command=compare_files, font=("Arial", 12))
    compare_button.grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()
