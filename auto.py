import win32com.client as win32
from time import sleep
import pandas as pd
import os

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule('FilePathCheckDLL', "SecurityModule")

excel = pd.read_excel("./database.xlsx", sheet_name='a-관리').fillna('')
hwp.Open("C:/Users/micke/Desktop/pythonProject/template/2022-1학기 간호관리학임상실습 임상실습현장지도자 평가표.hwp")
hwp.GetFieldList()
hwp.Run('SelectAll')
hwp.Run('Copy')
hwp.MovePos(3)

field_list = [i for i in hwp.GetFieldList().split('\x02')]
field_list

for i in range(len(excel)-1):
    hwp.MovePos(3)
    hwp.Run('Paste')
    hwp.MovePos(3)
for page in range(len(excel)):
    for field in field_list:
        if field == "사진":
            hwp.MoveToField(f'{field}{{{{{page}}}}}')
            base_path = "C:/Users/micke/Desktop/pythonProject/photo/" + str(excel[field].iloc[page]) + ".jpg"
            hwp.InsertPicture(base_path, Embedded=True, sizeoption=3)
            # hwp.InsertPicture("C:/Users/micke/Desktop/pythonProject/test.png", Embedded=True, sizeoption=3)
            hwp.Run("Close")
        else:
            hwp.PutFieldText(f'{field}{{{{{page}}}}}', excel[field].iloc[page])