import time
import win32com.client
import pandas as pd 
def atualizar_planilha(file):
    print('iniciar')
    app = win32com.client.Dispatch("Excel.Application")
    app.visible = 0
    workbook = app.Workbooks.open(file)
    workbook.RefreshAll()
    time.sleep(10)
    workbook.Save()
    app.Quit()
    exit
dir = r"C:\Users\Matheus\OneDrive\Nuvem\AMBIENTE_DE_DESENVOLVIMENTO\PYTHON\atualizador_de_dados_externos\config.csv"
lfile = pd.read_csv(dir,header=None)
for i in lfile.values:
    for v in i:
        atualizar_planilha(lfile)
