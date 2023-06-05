import time
import win32com.client as app
print('iniciar')
file = app.Dispatch("Excel.Application")
file.visible = 0
dirfile = r"C:\Users\Matheus\Downloads\teste.xlsx"
workbook = file.Workbooks.open(dirfile)
workbook.RefreshAll()
time.sleep(10)
workbook.Save()
file.Quit()
exit