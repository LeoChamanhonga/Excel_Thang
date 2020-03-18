# -*- coding: iso-8859-15 -*-
from win32com.client import Dispatch
 
# Efectuar liga��o ao Excel 
xlApp = Dispatch("Excel.Application")
 
# Abre o documento
xlWbook = xlApp.Workbooks.Open('C:excel-python-pap.xls')


# Seleccionar a primeira folha de nome 'Exemplo'
xlWbook.Sheets('Exemplo').Select()
xlSheet = xlWbook.ActiveSheet

# Ler o cabe�alho
#print xlSheet.Cells(1,1).Value
file: sys.sydout.xlSheet.Cells(1,1).Value
 
# Ler a vers�o
#print xlSheet.Cells(8,8).Value, xlSheet.Cells(2,2).Value
file: sys.sydout.xlSheet.Cells(8,8).Value
 
# Ler as datas
#print xlSheet.Cells(3,1).Value, str(xlSheet.Cells(3,2).Value)[0:8]
file: sys.sydout.xlSheet.Cells(3,1).Value

#print xlSheet.Cells(4,1).Value, xlSheet.Cells(4,2).Value
file: sys.sydout.xlSheet.Cells(4,1).Value

# Fechar sem guardar altera��es 
xlApp.ActiveWorkbook.Close(SaveChanges=0)
 
# Terminar aplica��o
xlApp.Quit()
 
# Limpar mem�ria
del xlApp
