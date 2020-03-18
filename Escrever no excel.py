# -*- coding: iso-8859-15 -*-
from win32com.client import Dispatch
import datetime
import sys
 
# Efectuar ligação ao Excel 
xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1

# Adicionar um novo documento
xlWbook = xlApp.Workbooks.Add()
 
# Eliminar todas as folhas excepto a primeira
for i in range(xlWbook.Sheets.Count, 1, -1):    
  xlWbook.Sheets(i).Delete()
 
# Seleccionar a primeira folha
xlWbook.Sheets(1).Select()
xlSheet = xlWbook.ActiveSheet
 
# Alterar o nome da folha para "Exemplo"
xlSheet.Name = "Exemplo"

#Pede nome ao utilizador
Nome = input('Insira seu Nome: ')


# Atribuir strings
xlSheet.Cells(1,1).Value = 'Criacao de Linhas No Excel'
xlSheet.Cells(2,1).Value = 'Python'
xlSheet.Cells(3,1).Value = 'Criada Em'
xlSheet.Cells(4,1).Value = 'Modificada Em'
#xlSheet.Cells(5,1).Value = 'Mas 2'
xlSheet.Cells(7,6).Value = 'Criado Por: '
xlSheet.Cells(7,8).Value = str (Nome) 
xlSheet.Cells(8,6).Value = 'Versao Do Python'
 
# Atribuir valor numérico
xlSheet.Cells(8,8).Value = str(sys.version_info[0]) + "." + str(sys.version_info[1])
 
# Atribuir uma data
#xlSheet.Cells(3,2).Value = str (datetime.date.today)
 
# Atribuir uma fórmula / Escreve a Data e Hora da Alteracao
xlSheet.Cells(4,2).Formula = '=Now()'

# Toque artistico para o título :)
xlSheet.Cells(1,1).Font.Bold = True
xlSheet.Cells(1,1).Font.ColorIndex = 5
xlSheet.Range("A1:B1").Merge()

# Campo do Criado
xlSheet.Cells(7,8).Font.Bold = True
xlSheet.Cells(7,8).Font.ColorIndex = 7



#Acrescimos JUNTAR COLUNAS
xlSheet.Range("L13:M13").Merge()
xlSheet.Range("F7:G7").Merge()
 
# Ajustar tamanhos das colunas
for i in range(1,3):
  xlSheet.Columns(i).Autofit

# Gravar o documento
xlWbook.SaveAs ('C:excel-python-pap.xls')
#xlWbook.SaveAs ('C:\Users\Leonel K4T\OneDrive\deV\Py\Excel Thang excel-python.xls')
 
# Terminar aplicação
xlApp.Quit()
 
# Limpar memória
del xlApp
