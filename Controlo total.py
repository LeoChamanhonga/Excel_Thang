# -*- coding: iso-8859-15 -*-
from win32com.client import Dispatch
 
# Efectuar liga��o ao Excel 
xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1
 
# Abre o documento
xlWbook = xlApp.Workbooks.Open('C:excel-python-pap.xls')


# Seleccionar a folha de nome 'Exemplo'
xlWbook.Sheets('Exemplo').Select()
xlSheet = xlWbook.ActiveSheet

# Seleccionar a informa��o a imprimir
xlSheet.PageSetup.PrintArea = "A1:B4"
 
# Imprimir as c�lulas seleccionadas
xlSheet.PrintOut()

# Falar
xlApp.Speech.Speak("Python rules!")

# Fechar sem guardar altera��es 
xlApp.ActiveWorkbook.Close(SaveChanges=0)
 
# Terminar aplica��o
xlApp.Quit()
 
# Limpar mem�ria
del xlApp
