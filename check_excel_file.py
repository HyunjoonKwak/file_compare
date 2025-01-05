import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

# 보완: 셀 값 비교 함수
def are_values_equal(val1, val2):
    """
    값이 모두 None이면 동일하다고 간주,
    그렇지 않으면 str로 변환하여 비교합니다.
    """
    if val1 is None and val2 is None:
        return True
    if val1 is None or val2 is None:
        return False
    return str(val1) == str(val2)

def highlight_changes(file1, file2, output_file):
    """
    파일 경로 file1, file2를 비교한 뒤, 결과를 output_file로 저장합니다.
    - df1에만 있는 시트 => 전체가 '삭제된' 것으로 간주 (노란색)
    - df2에만 있는 시트 => 전체가 '추가된' 것으로 간주 (파란색)
    - 공통 시트 => 셀 단위로 변경/추가/삭제 여부 비교
    """
    # 파일 읽기
    df1 = pd.read_excel(file1, sheet_name=None)
    df2 = pd.read_excel(file2, sheet_name=None)
    
    # Output Workbook 생성
    wb = Workbook()
    # 기본으로 생성되는 워크시트 제거 (sheet 이름이 "Sheet"인 기본 워크시트)
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # 색상 정의
    # (start_color/end_color 동일하게 설정)
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")   # 변경된 부분
    blue_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")  # 추가된 부분
    yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")# 삭제된 부분

    # 1) 공통 시트 비교
    common_sheets = set(df1.keys()).intersection(set(df2.keys()))
    
    for sheet_name in common_sheets:
        ws = wb.create_sheet(title=sheet_name)
        df1_sheet = df1[sheet_name]
        df2_sheet = df2[sheet_name]

        max_rows = max(len(df1_sheet), len(df2_sheet))
        max_cols = max(len(df1_sheet.columns), len(df2_sheet.columns))

        for i in range(max_rows):
            for j in range(max_cols):
                # df1, df2에서 i, j 셀 값 추출
                cell_value1 = None
                if i < len(df1_sheet.index) and j < len(df1_sheet.columns):
                    cell_value1 = df1_sheet.iloc[i, j]
                
                cell_value2 = None
                if i < len(df2_sheet.index) and j < len(df2_sheet.columns):
                    cell_value2 = df2_sheet.iloc[i, j]
                
                # 결과 시트 셀
                cell = ws.cell(row=i+1, column=j+1)
                
                # 우선 cell.value는 df2의 값이 있으면 그것을 사용, 아니면 df1 값
                cell.value = cell_value2 if cell_value2 is not None else cell_value1

                # 비교 로직
                if not are_values_equal(cell_value1, cell_value2):
                    # 변경된 부분 (df1, df2 둘 다 값이 존재하나 서로 다를 때)
                    if cell_value1 is not None and cell_value2 is not None:
                        cell.fill = red_fill
                    # 추가된 부분 (df1은 None, df2에 값이 있을 때)
                    elif cell_value1 is None and cell_value2 is not None:
                        cell.fill = blue_fill
                    # 삭제된 부분 (df1에 값이 있고, df2는 None일 때)
                    elif cell_value1 is not None and cell_value2 is None:
                        cell.fill = yellow_fill

    # 2) df1에만 존재하는 시트 => 모두 '삭제된' 것으로 간주
    df1_only_sheets = set(df1.keys()) - set(df2.keys())
    for sheet_name in df1_only_sheets:
        ws = wb.create_sheet(title=f"{sheet_name}_DELETED")  # 어떤 식으로 명명할지 자유
        df1_sheet = df1[sheet_name]
        
        for row_idx, row_data in enumerate(df1_sheet.itertuples(index=False, name=None)):
            ws.append(row_data)
        
        # 추가로 색칠
        max_rows = len(df1_sheet.index)
        max_cols = len(df1_sheet.columns)
        for i in range(max_rows):
            for j in range(max_cols):
                cell = ws.cell(row=i+1, column=j+1)
                cell.fill = yellow_fill  # 모두 삭제된 것으로 표시

    # 3) df2에만 존재하는 시트 => 모두 '추가된' 것으로 간주
    df2_only_sheets = set(df2.keys()) - set(df1.keys())
    for sheet_name in df2_only_sheets:
        ws = wb.create_sheet(title=f"{sheet_name}_ADDED")
        df2_sheet = df2[sheet_name]

        for row_data in df2_sheet.itertuples(index=False, name=None):
            ws.append(row_data)
        
        # 추가로 색칠
        max_rows = len(df2_sheet.index)
        max_cols = len(df2_sheet.columns)
        for i in range(max_rows):
            for j in range(max_cols):
                cell = ws.cell(row=i+1, column=j+1)
                cell.fill = blue_fill  # 모두 추가된 것으로 표시

    # 4) 새로운 파일 저장
    wb.save(output_file)


# ────────────── #
#  GUI 인터페이스  #
# ────────────── #

def select_first_file():
    file = filedialog.askopenfilename(title="Select the first Excel file", filetypes=[("Excel files", "*.xlsx")])
    if file:
        first_file_var.set(file)

def select_second_file():
    file = filedialog.askopenfilename(title="Select the second Excel file", filetypes=[("Excel files", "*.xlsx")])
    if file:
        second_file_var.set(file)

def compare_files():
    file1 = first_file_var.get()
    file2 = second_file_var.get()
    
    if not file1 or not file2:
        messagebox.showerror("Error", "Both files must be selected!")
        return

    output_file = filedialog.asksaveasfilename(title="Save the comparison result", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file:
        return

    try:
        highlight_changes(file1, file2, output_file)
        messagebox.showinfo("Success", f"Comparison complete! Result saved to:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel File Comparison Tool (Enhanced)")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack()

    # 첫 번째 파일 선택
    tk.Label(frame, text="First Excel File:", font=("Arial", 12)).grid(row=0, column=0, sticky="e", pady=5)
    first_file_var = tk.StringVar()
    tk.Entry(frame, textvariable=first_file_var, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
    tk.Button(frame, text="Browse", command=select_first_file).grid(row=0, column=2, padx=5, pady=5)

    # 두 번째 파일 선택
    tk.Label(frame, text="Second Excel File:", font=("Arial", 12)).grid(row=1, column=0, sticky="e", pady=5)
    second_file_var = tk.StringVar()
    tk.Entry(frame, textvariable=second_file_var, width=50, state="readonly").grid(row=1, column=1, padx=5, pady=5)
    tk.Button(frame, text="Browse", command=select_second_file).grid(row=1, column=2, padx=5, pady=5)

    # 비교 실행 버튼
    compare_button = tk.Button(frame, text="Compare Files", command=compare_files, font=("Arial", 12))
    compare_button.grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()
