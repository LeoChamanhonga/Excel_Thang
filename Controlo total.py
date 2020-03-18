# -*- coding: iso-8859-15 -*-
from win32com.client import Dispatch
 
# Efectuar ligação ao Excel 
xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1
 
# Abre o documento
xlWbook = xlApp.Workbooks.Open('C:excel-python-pap.xls')


# Seleccionar a folha de nome 'Exemplo'
xlWbook.Sheets('Exemplo').Select()
xlSheet = xlWbook.ActiveSheet

# Seleccionar a informação a imprimir
xlSheet.PageSetup.PrintArea = "A1:B4"
 
# Imprimir as células seleccionadas
xlSheet.PrintOut()

# Falar
xlApp.Speech.Speak("Python rules!")

# Fechar sem guardar alterações 
xlApp.ActiveWorkbook.Close(SaveChanges=0)
 
# Terminar aplicação
xlApp.Quit()
 
# Limpar memória
del xlApp
