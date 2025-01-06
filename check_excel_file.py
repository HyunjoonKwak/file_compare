import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 비교에서 제외할 시트
EXCLUDED_SHEETS = {"Cover", "test description", "Results"}

# 색상 정의
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 변경(빨간색)
BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")   # 추가(파란색)
YELLOW_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # 삭제(노란색)

# 고유 키 컬럼 (열 이름)
TEST_ID_COLUMN = "Test ID"

def load_sheets(file_path):
    """
    주어진 파일 경로(file_path)를 열어,
    1) EXCLUDED_SHEETS (Cover, test description, Results)는 건너뛰고
    2) History 시트는 skiprows=2 (즉, 3행이 헤더)
    3) 그 외 시트는 skiprows=4 (즉, 5행이 헤더)
    로 읽어서 딕셔너리에 담아 리턴합니다.
    dict[sheet_name] = 해당 시트의 DataFrame
    """
    # 먼저 파일에 어떤 시트들이 있는지 확인
    all_sheets = pd.ExcelFile(file_path).sheet_names
    
    sheet_dict = {}
    for sheet_name in all_sheets:
        # 1) 제외 목록인 경우 건너뛴다
        if sheet_name in EXCLUDED_SHEETS:
            continue
        
        # 2) History 시트인지 판별
        if sheet_name == "History":
            # 3행이 헤더라고 가정 -> skiprows=2
            # (row 1, row 2를 건너뛰고, row 3을 헤더로 사용)
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
        else:
            # 그 외 시트 -> 5행이 헤더라고 가정 -> skiprows=4
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=4)
        
        sheet_dict[sheet_name] = df
    
    return sheet_dict


def compare_two_sheets_with_key(sheet_name, df1_sheet, df2_sheet, wb, key_column=TEST_ID_COLUMN):
    """
    'Test ID' 열을 기준으로 df1_sheet, df2_sheet를 비교해
    - left_only => 삭제(노란색)
    - right_only => 추가(파란색)
    - both인데 old != new => 변경(빨간색)
    결과를 wb(Workbook)의 새 시트(sheet_name)에 기록합니다.
    """
    merged_df = pd.merge(
        df1_sheet, df2_sheet,
        on=key_column, how="outer", indicator=True,
        suffixes=("_old", "_new")
    )

    ws = wb.create_sheet(title=sheet_name)

    df1_cols = list(df1_sheet.columns)
    df2_cols = list(df2_sheet.columns)

    # key_column("Test ID")는 앞으로 빼고, 나머지는 old/new 순으로 정렬
    df1_cols.remove(key_column)  # ['OtherCol1', 'OtherCol2', ...]
    df2_cols.remove(key_column)  # ['OtherCol1', 'OtherCol2', ...]

    output_cols = [key_column]
    for col in df1_cols:
        output_cols.append(col + "_old")
    for col in df2_cols:
        output_cols.append(col + "_new")
    output_cols.append("_merge")

    # 실제로 merged_df에 존재하는 컬럼만 사용
    final_cols = [c for c in output_cols if c in merged_df.columns]

    # 헤더 추가
    ws.append(final_cols)

    # 데이터 추가
    for _, row_data in merged_df[final_cols].iterrows():
        ws.append(row_data.tolist())

    # 색상 처리
    merge_col_idx = final_cols.index("_merge") + 1  # openpyxl 1-based
    old_map = {}
    new_map = {}

    # old/new 컬럼 인덱스 매핑
    for idx, col_name in enumerate(final_cols, start=1):
        if col_name.endswith("_old"):
            original_col = col_name[:-4]  # "Name_old" -> "Name"
            old_map[original_col] = idx
        elif col_name.endswith("_new"):
            original_col = col_name[:-4]
            new_map[original_col] = idx

    max_row = ws.max_row
    max_col = ws.max_column

    for row_idx in range(2, max_row + 1):
        merge_value = ws.cell(row=row_idx, column=merge_col_idx).value

        if merge_value == "left_only":
            # df1에만 있음 => 삭제
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = YELLOW_FILL

        elif merge_value == "right_only":
            # df2에만 있음 => 추가
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = BLUE_FILL

        else:
            # both
            # old/new 비교
            for orig_col, old_idx in old_map.items():
                new_idx = new_map.get(orig_col)
                if not new_idx:
                    continue
                val_old = ws.cell(row=row_idx, column=old_idx).value
                val_new = ws.cell(row=row_idx, column=new_idx).value

                if val_old != val_new:
                    ws.cell(row=row_idx, column=old_idx).fill = RED_FILL
                    ws.cell(row=row_idx, column=new_idx).fill = RED_FILL


def highlight_changes_with_test_id(file1, file2, output_file):
    """
    - file1, file2 경로의 Excel 파일을 열고,
    - History 시트는 3행이 헤더, 나머지 시트는 5행이 헤더로 간주해 로드
    - 'Test ID' 컬럼을 기준으로 비교(행 매칭)
    - EXCLUDED_SHEETS는 건너뛴다
    - 결과를 output_file에 저장
    """
    # 1) 파일별 시트 로드 (skiprows 다르게 적용)
    df1_dict = load_sheets(file1)  # {sheet_name: df}
    df2_dict = load_sheets(file2)

    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    # 시트 집합
    df1_sheets = set(df1_dict.keys())
    df2_sheets = set(df2_dict.keys())

    # 공통 시트
    common_sheets = df1_sheets & df2_sheets

    for sheet_name in common_sheets:
        df1_sheet = df1_dict[sheet_name]
        df2_sheet = df2_dict[sheet_name]

        # 만약 'Test ID'가 없으면 비교 불가
        if TEST_ID_COLUMN not in df1_sheet.columns or TEST_ID_COLUMN not in df2_sheet.columns:
            ws = wb.create_sheet(title=f"{sheet_name}_NO_KEY")
            ws.cell(row=1, column=1).value = f"'{TEST_ID_COLUMN}' 열이 없어 비교 불가"
            continue

        compare_two_sheets_with_key(sheet_name, df1_sheet, df2_sheet, wb, TEST_ID_COLUMN)

    # df1에만 있는 시트 => _DELETED
    df1_only = df1_sheets - df2_sheets
    for sheet_name in df1_only:
        df_sheet = df1_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")
        for row_data in df_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 모두 노란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL

    # df2에만 있는 시트 => _ADDED
    df2_only = df2_sheets - df1_sheets
    for sheet_name in df2_only:
        df_sheet = df2_dict[sheet_name]
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        for row_data in df_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        # 모두 파란색
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = BLUE_FILL

    wb.save(output_file)

# ────────────── #
#  GUI 인터페이스  #
# ────────────── #

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
        highlight_changes_with_test_id(file1, file2, output_file)
        messagebox.showinfo("Success", f"Comparison complete!\nResult saved to: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel File Comparison (History=C3, Others=B5, Key='Test ID')")

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
