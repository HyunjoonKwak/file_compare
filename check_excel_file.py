import openpyxl
from openpyxl.styles import PatternFill
from tkinter import Tk, filedialog, messagebox, Button, Label, Entry

def compare_excel_files(file1, file2, output_file):
    # 파일 불러오기
    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)

    # 결과 파일 생성
    wb_result = openpyxl.Workbook()
    wb_result.remove(wb_result.active)  # 기본 시트 제거

    # 색상 정의
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # 추가된 행
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # 삭제된 행
    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # 변경된 셀

    for sheet_name in wb1.sheetnames:
        if sheet_name in ["Cover", "test description", "Results"]:
            continue

        ws1 = wb1[sheet_name]
        ws2 = wb2[sheet_name]
        ws_result = wb_result.create_sheet(sheet_name)

        if sheet_name == "History":
            # History 시트 처리 (4열부터 비교)
            rows1 = list(ws1.iter_rows(min_row=2, min_col=4, values_only=True))
            rows2 = list(ws2.iter_rows(min_row=2, min_col=4, values_only=True))

            ws_result.append(["(추가된 행)"] + [f"열 {i+4}" for i in range(len(rows2[0]))])  # 헤더
            for row in rows2:
                if row not in rows1:
                    ws_result.append(list(row))

        else:
            # 나머지 시트 처리 (B열 기준)
            rows1 = {row[1]: row for row in ws1.iter_rows(min_row=2, values_only=True) if row[1] and str(row[1]).startswith("IO")}
            rows2 = {row[1]: row for row in ws2.iter_rows(min_row=2, values_only=True) if row[1] and str(row[1]).startswith("IO")}

            keys1 = set(rows1.keys())
            keys2 = set(rows2.keys())

            added_keys = keys2 - keys1
            deleted_keys = keys1 - keys2
            common_keys = keys1 & keys2

            # 추가된 행
            ws_result.append(["(추가된 행)"] + [f"열 {i+1}" for i in range(len(next(iter(rows2.values()))))])  # 헤더
            for key in added_keys:
                ws_result.append(rows2[key])
                for cell in ws_result[ws_result.max_row]:
                    cell.fill = red_fill

            # 삭제된 행
            ws_result.append(["(삭제된 행)"] + [f"열 {i+1}" for i in range(len(next(iter(rows1.values()))))])  # 헤더
            for key in deleted_keys:
                ws_result.append(rows1[key])
                for cell in ws_result[ws_result.max_row]:
                    cell.fill = gray_fill

            # 변경된 행
            ws_result.append(["(변경된 행)"] + [f"열 {i+1}" for i in range(len(next(iter(rows1.values()))))])  # 헤더
            for key in common_keys:
                row1 = rows1[key]
                row2 = rows2[key]
                new_row = []
                for cell1, cell2 in zip(row1, row2):
                    if cell1 == cell2:
                        new_row.append(cell1)
                    else:
                        new_row.append(cell2)
                ws_result.append(new_row)
                for i, (cell1, cell2) in enumerate(zip(row1, row2)):
                    if cell1 != cell2:
                        ws_result.cell(row=ws_result.max_row, column=i + 1).fill = blue_fill

    # 파일 확장자 확인 및 저장
    if not output_file.endswith(".xlsx"):
        output_file += ".xlsx"

    wb_result.save(output_file)
    messagebox.showinfo("완료", f"비교 결과가 {output_file}에 저장되었습니다.")

def select_file(label):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        label.config(text=filename)
    return filename

def start_comparison():
    file1 = file1_label.cget("text")
    file2 = file2_label.cget("text")
    output_file = output_entry.get()

    if not file1 or not file2 or not output_file:
        messagebox.showerror("오류", "모든 파일과 저장 경로를 지정해주세요.")
        return

    try:
        compare_excel_files(file1, file2, output_file)
    except Exception as e:
        messagebox.showerror("오류", f"파일 비교 중 오류가 발생했습니다: {e}")

# Tkinter 인터페이스 생성
root = Tk()
root.title("Excel 파일 비교")

# 파일 선택 UI
Label(root, text="첫 번째 파일:").grid(row=0, column=0, padx=10, pady=5)
file1_label = Label(root, text="", width=50, anchor="w", relief="solid")
file1_label.grid(row=0, column=1, padx=10, pady=5)
Button(root, text="파일 선택", command=lambda: select_file(file1_label)).grid(row=0, column=2, padx=10, pady=5)

Label(root, text="두 번째 파일:").grid(row=1, column=0, padx=10, pady=5)
file2_label = Label(root, text="", width=50, anchor="w", relief="solid")
file2_label.grid(row=1, column=1, padx=10, pady=5)
Button(root, text="파일 선택", command=lambda: select_file(file2_label)).grid(row=1, column=2, padx=10, pady=5)

# 결과 저장 파일명 입력
Label(root, text="저장 파일명:").grid(row=2, column=0, padx=10, pady=5)
output_entry = Entry(root, width=53)
output_entry.grid(row=2, column=1, padx=10, pady=5)

# 비교 시작 버튼
Button(root, text="비교 시작", command=start_comparison).grid(row=3, column=0, columnspan=3, pady=20)

root.mainloop()
