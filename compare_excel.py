    
## 파일 불러오기
from IPython.display import display
from ipyfilechooser import FileChooser
import ipywidgets as widgets
import os
from tkinter import filedialog
from tkinter import messagebox

import pandas as pd    

def compare_excel(old_xlsx, new_xlsx, column_name):
    df_old = pd.read_excel(old_xlsx)
    df_new = pd.read_excel(new_xlsx)

    # 불러온 데이터의 버전 구분
    df_old['ver'] = 'old'
    df_new['ver'] = 'new'

    id_dropped = set(df_old[column_name]) - set(df_new[column_name])
    id_added = set(df_new[column_name]) - set(df_old[column_name])

    # 삭제된 데이터
    df_dropped = df_old[df_old[column_name].isin(id_dropped)].iloc[:,:-1]
    # 추가된 데이터
    df_added = df_new[df_new[column_name].isin(id_added)].iloc[:,:-1]

    df_concatted = pd.concat([df_old, df_new], ignore_index=True)
    changes = df_concatted.drop_duplicates(df_concatted.columns[:-1], keep='last')
    duplicated_list = changes[changes[column_name].duplicated()][column_name].to_list()
    df_changed = changes[changes[column_name].isin(duplicated_list)]

    df_changed_old = df_changed[df_changed['ver'] == 'old'].iloc[:,:-1]
    df_changed_old.sort_values(by=column_name, inplace=True)

    df_changed_new = df_changed[df_changed['ver'] == 'new'].iloc[:,:-1]
    df_changed_new.sort_values(by=column_name, inplace=True)

    # 정보가 변경된 데이터 정리
    df_info_changed = df_changed_old.copy()
    for i in range(len(df_changed_new.index)):
        for j in range(len(df_changed_new.columns)):
            if (df_changed_new.iloc[i, j] != df_changed_old.iloc[i, j]):
                df_info_changed.iloc[i,j] = str(df_changed_old.iloc[i, j]) + " ==> " + str(df_changed_new.iloc[i,j])

    # 엑셀 저장            
    with pd.ExcelWriter('compared_result.xlsx') as writer:
        df_info_changed.to_excel(writer, sheet_name='info changed', index=False)
        df_added.to_excel(writer, sheet_name='added', index=False)
        df_dropped.to_excel(writer, sheet_name='dropped', index=False)              

###############################################################################################################
#   불러올 파일 선택
#   Select 버튼으로 불러올 파일을 선택한다. (*.xlsx)
###############################################################################################################
old_file = []                                          #파일 목록 담을 리스트 생성
new_file = []
old_files = filedialog.askopenfilenames(initialdir="/",title = "Old 파일을 선택 해 주세요", filetypes = (("*.xlsx","*xlsx"),("*.xls","*xls"),("*.csv","*csv")))
#files 변수에 선택 파일 경로 넣기
old_file = old_files[0]

new_files = filedialog.askopenfilenames(initialdir="/",title = "New 파일을 선택 해 주세요", filetypes = (("*.xlsx","*xlsx"),("*.xls","*xls"),("*.csv","*csv")))
#files 변수에 선택 파일 경로 넣기
new_file = new_files[0]


if old_files == '':
    messagebox.showwarning("경고", "파일을 추가 하세요")    #파일 선택 안했을 때 메세지 출력

old_file = old_file.replace(")","").replace("(","").replace(",","")
new_file = new_file.replace(")","").replace("(","").replace(",","")

print("old file path = ",old_file)    #files 리스트 값 출력
print("new file path = ",new_file)    #files 리스트 값 출력

compare_excel(old_file, new_file, "NO")