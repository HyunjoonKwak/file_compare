import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 비교 제외할 시트
EXCLUDED_SHEETS = {"Cover", "test description", "Results"}

# 색상 정의
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 변경된 부분
BLUE_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")   # 추가된(행) 부분
YELLOW_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # 삭제된(행) 부분


def compare_two_sheets_with_key(sheet_name, df1_sheet, df2_sheet, wb, key_column="ID"):
    """
    고유 키(기본값 'ID')를 기준으로 df1_sheet와 df2_sheet를 비교한 뒤,
    결과를 wb(Workbook)에 sheet_name으로 시트를 만들어 기록한다.
    """
    # (1) 두 시트를 key_column 기준으로 outer join
    #     suffixes=('_old','_new') 를 붙여서, df1의 컬럼과 df2의 컬럼이 겹치지 않게 설정
    merged_df = pd.merge(
        df1_sheet, df2_sheet,
        on=key_column, how="outer", indicator=True,
        suffixes=("_old", "_new")
    )

    # 시트 생성
    ws = wb.create_sheet(title=sheet_name)

    # (2) 엑셀에 기록할 컬럼 순서를 정한다.
    #     df1의 컬럼, df2의 컬럼 중에서 key_column은 맨 앞으로 빼고,
    #     key_column을 제외한 나머지는 df1_old, df2_new 순서로.
    # 먼저 df1_sheet, df2_sheet의 컬럼 목록을 확인
    df1_cols = list(df1_sheet.columns)
    df2_cols = list(df2_sheet.columns)

    # key_column을 맨 앞으로 두고, 나머지를 old -> new 순으로 배치
    # df1에서는 old suffix, df2에서는 new suffix가 붙을 것
    df1_cols.remove(key_column)
    df2_cols.remove(key_column)

    # 최종 출력할 컬럼 (key_column / df1_old / df2_new / 그리고 indicator('_merge'))
    output_cols = [key_column]
    for col in df1_cols:
        output_cols.append(col + "_old")
    for col in df2_cols:
        output_cols.append(col + "_new")
    output_cols.append("_merge")  # 행이 left-only/ right-only/ both 인지 알려주는 컬럼

    # 실제로 merged_df에 존재하는 컬럼 중에서 output_cols 순서를 지키되, 없는 컬럼은 건너뛴다.
    final_cols = [c for c in output_cols if c in merged_df.columns]

    # (3) 엑셀 시트에 헤더 기록
    ws.append(final_cols)

    # (4) 각 행을 순회하며 값 채우기 + 색상
    for row_idx, row_data in merged_df[final_cols].iterrows():
        # row_data는 시리즈 형태, 이를 리스트로 변환
        row_list = row_data.tolist()
        ws.append(row_list)

    # (5) 색상 하이라이트 로직
    #     - _merge 컬럼이 'left_only' => df1에는 있고, df2에는 없는 행 => 삭제(노란색)
    #     - _merge 컬럼이 'right_only' => df2에만 있는 행 => 추가(파란색)
    #     - 'both' 인 경우, 셀 단위로 비교
    #       => 예) A_old != A_new 인 경우 빨간색
    max_row = ws.max_row
    max_col = ws.max_column

    # _merge 컬럼의 인덱스를 찾는다 (final_cols에서 _merge가 몇 번째인지)
    merge_col_index = final_cols.index("_merge") + 1  # +1은 openpyxl의 1-based index

    # df1_old / df2_new 컬럼이 어디 있는지 찾기
    # 예를 들어 "Name_old", "Name_new", "Age_old", "Age_new" 등
    old_columns_map = {}  # { original_col_name: col_index_in_worksheet_for_old }
    new_columns_map = {}  # { original_col_name: col_index_in_worksheet_for_new }

    for c_idx, col_name in enumerate(final_cols, start=1):
        # col_name 예) "Name_old"
        if col_name.endswith("_old"):
            old_columns_map[col_name[:-4]] = c_idx  # _old 제거
        elif col_name.endswith("_new"):
            new_columns_map[col_name[:-4]] = c_idx  # _new 제거

    for row_idx in range(2, max_row + 1):  # 2행부터 데이터 존재
        merge_value = ws.cell(row=row_idx, column=merge_col_index).value

        if merge_value == "left_only":
            # df1에만 존재 => 삭제된 행
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = YELLOW_FILL

        elif merge_value == "right_only":
            # df2에만 존재 => 추가된 행
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = BLUE_FILL

        else:
            # both => 셀 단위로 비교 (xxx_old != xxx_new)
            # old_columns_map, new_columns_map 사용
            for orig_col in old_columns_map.keys():
                old_col_idx = old_columns_map[orig_col]
                new_col_idx = new_columns_map.get(orig_col)

                # 만약 df2에 해당 컬럼이 없으면 비교 불가능
                if not new_col_idx:
                    continue

                val_old = ws.cell(row=row_idx, column=old_col_idx).value
                val_new = ws.cell(row=row_idx, column=new_col_idx).value

                # 값이 다를 경우만 빨간색 표시
                if val_old != val_new:
                    ws.cell(row=row_idx, column=old_col_idx).fill = RED_FILL
                    ws.cell(row=row_idx, column=new_col_idx).fill = RED_FILL


def highlight_changes(file1, file2, output_file, key_column="ID"):
    """
    file1, file2를 비교해서 output_file로 결과 저장.
    - key_column(기본값 "ID")을 기준으로 행 매칭
    - 제외 시트(Cover, test description, Results)는 건너뜀
    - df1에만 있는 시트 => _DELETED 시트 (노란색)
    - df2에만 있는 시트 => _ADDED 시트 (파란색)
    - 공통 시트 => key_column을 기준으로 병합 후 셀 단위 비교
    """
    df1 = pd.read_excel(file1, sheet_name=None)
    df2 = pd.read_excel(file2, sheet_name=None)
    
    wb = Workbook()
    # 기본 생성되는 시트 제거
    wb.remove(wb.active)

    df1_sheets = set(df1.keys())
    df2_sheets = set(df2.keys())

    # 비교 대상 시트: 공통 시트 중에서 제외 목록에 없는 시트만
    common_sheets = (df1_sheets & df2_sheets) - EXCLUDED_SHEETS

    for sheet_name in common_sheets:
        # key_column이 없으면 KeyError가 날 수 있으므로 예외처리
        if key_column not in df1[sheet_name].columns or key_column not in df2[sheet_name].columns:
            # 키 컬럼이 없는 시트는 그냥 옛날 방식으로 처리하거나, 스킵할 수도 있음
            # 여기서는 "키 컬럼 없음" 시트라고 제목 붙여서 생성
            ws = wb.create_sheet(title=f"{sheet_name}_NO_KEY")
            ws.cell(row=1, column=1, value="해당 시트에는 키 컬럼이 없어 비교 불가")
            continue

        df1_sheet = df1[sheet_name]
        df2_sheet = df2[sheet_name]

        compare_two_sheets_with_key(sheet_name, df1_sheet, df2_sheet, wb, key_column)

    # df1에만 있고 df2에는 없는 시트(비교 제외 목록은 제외)
    df1_only = (df1_sheets - df2_sheets) - EXCLUDED_SHEETS
    for sheet_name in df1_only:
        # 통째로 "삭제됨" 표시
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")
        sheet_df = df1[sheet_name]
        for row_data in sheet_df.itertuples(index=False, name=None):
            ws.append(row_data)
        max_r = ws.max_row
        max_c = ws.max_column
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ws.cell(row=r, column=c).fill = YELLOW_FILL

    # df2에만 있고 df1에는 없는 시트(비교 제외 목록은 제외)
    df2_only = (df2_sheets - df1_sheets) - EXCLUDED_SHEETS
    for sheet_name in df2_only:
        # 통째로 "추가됨" 표시
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        sheet_df = df2[sheet_name]
        for row_data in sheet_df.itertuples(index=False, name=None):
            ws.append(row_data)
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

    # 키 컬럼 명을 사용자 입력으로 받는다고 가정(기본값 'ID')
    key_col = key_column_var.get().strip()
    if not key_col:
        key_col = "ID"

    output_file = filedialog.asksaveasfilename(
        title="Save the comparison result",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_file:
        return

    try:
        highlight_changes(file1, file2, output_file, key_column=key_col)
        messagebox.showinfo("Success", f"Comparison complete!\nResult saved to: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel File Comparison Tool (Row insertion handled by Key)")

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

    tk.Label(frame, text="Key Column Name:", font=("Arial", 12)).grid(row=2, column=0, sticky="e", pady=5)
    key_column_var = tk.StringVar(value="ID")  # 기본값 ID
    tk.Entry(frame, textvariable=key_column_var, width=20).grid(row=2, column=1, padx=5, pady=5, sticky="w")

    compare_button = tk.Button(frame, text="Compare Files", command=compare_files, font=("Arial", 12))
    compare_button.grid(row=3, column=0, columnspan=3, pady=20)

    root.mainloop()
