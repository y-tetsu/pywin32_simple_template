import win32com.client
from pathlib import Path


# Excelがすでに起動していたら閉じるよう警告する
try:
    if win32com.client.GetObject(Class='Excel.Application'):
        raise RuntimeError('Close all Excel applications!')

except win32com.client.pywintypes.com_error:
    pass


# 既存のExcelファイルを開いて何かする
file_path = Path('読み込みたいファイル.xlsx')
file_name = str(file_path.absolute())

try:
    app = win32com.client.Dispatch('Excel.Application')
    wb = app.Workbooks.Open(file_name)

    # ↓↓↓ 処理を書く ↓↓↓

    # ↑↑↑ 処理を書く ↑↑↑

finally:
    wb.Close()
    app.Quit()


# 新規にExcelファイルを作って何かする
file_path = Path('書き込みたいファイル.xlsx')
file_name = str(file_path.absolute())

try:
    app = win32com.client.Dispatch('Excel.Application')
    wb = app.Workbooks.Add()

    if file_path.is_file():
        raise RuntimeError(f'{file_name} already exists!')

    # ↓↓↓ 処理を書く ↓↓↓

    # ↑↑↑ 処理を書く ↑↑↑

    wb.SaveAs(Filename=file_name)

finally:
    wb.Close()
    app.Quit()
