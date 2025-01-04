import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox

def highlight_changes(file1, file2, output_file):
    # 파일 읽기
    df1 = pd.read_excel(file1, sheet_name=None)
    df2 = pd.read_excel(file2, sheet_name=None)
    
    # Output Workbook 생성
    wb = Workbook()
    
    # 색상 정의
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # 변경된 부분
    blue_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")  # 추가된 부분
    yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")  # 삭제된 부분

    # 시트별 비교
    for sheet_name in df1.keys():
        if sheet_name not in df2:
            continue  # df2에 없는 sheet는 무시
        
        ws = wb.create_sheet(title=sheet_name)
        df1_sheet = df1[sheet_name]
        df2_sheet = df2[sheet_name]

        max_rows = max(len(df1_sheet), len(df2_sheet))
        max_cols = max(len(df1_sheet.columns), len(df2_sheet.columns))

        for i in range(max_rows):
            for j in range(max_cols):
                cell_value1 = df1_sheet.iloc[i, j] if i < len(df1_sheet) and j < len(df1_sheet.columns) else None
                cell_value2 = df2_sheet.iloc[i, j] if i < len(df2_sheet) and j < len(df2_sheet.columns) else None

                cell = ws.cell(row=i+1, column=j+1)
                cell.value = cell_value2 if cell_value2 is not None else cell_value1

                # 변경된 부분
                if cell_value1 != cell_value2:
                    if cell_value1 is not None and cell_value2 is not None:
                        cell.fill = red_fill  # 변경된 부분
                    elif cell_value1 is None:
                        cell.fill = blue_fill  # 추가된 부분
                    elif cell_value2 is None:
                        cell.fill = yellow_fill  # 삭제된 부분

    # 남은 df2의 시트 추가
    for sheet_name in df2.keys():
        if sheet_name not in df1:
            ws = wb.create_sheet(title=sheet_name)
            for r in df2[sheet_name].itertuples(index=False, name=None):
                ws.append(r)

    # 새로운 파일 저장
    wb.save(output_file)

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

# GUI 생성
root = tk.Tk()
root.title("Excel File Comparison Tool")

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

# GUI 실행
root.mainloop()