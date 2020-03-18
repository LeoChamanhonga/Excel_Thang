# -*- coding: utf-8 -*-
# Script para gravar em planilha Excel

import os
import numpy as np
import pandas as pd
import win32com.client as win32
from xlsxwriter.utility import xl_range, xl_cell_to_rowcol

# Função para abrir arquivo Excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
def openWorkbook(xlfile):
	global excel
	global win32
	try:        
	    xlwb = win32.GetObject(xlfile)            
	except Exception as e:
	    try:
	        xlwb = excel.Workbooks.Open(xlfile)
	    except Exception as e:
	        print(e)
	        xlwb = None                    
	return(xlwb)

# Função para ler planilha e gravar em data frame
def dataFramexl(sheet, rng):
	array = np.array(sheet.Range(rng).Value)
	return pd.DataFrame( array[1:], columns=array[0] )

# Função para imprimir tabela
def printxl(sheet,start, array ):
	array = np.array(array)
	start = xl_cell_to_rowcol(start)
	sheet.Range( xl_range(
		start[0],
		start[1],
		len(array)+start[0]-1,
		len(array[0])+start[1]-1
		)
	).Value = array

# Função para imprimir linha
def printxl_line(sheet,start, array):
	array = np.array(array)
	start = xl_cell_to_rowcol(start)
	sheet.Range( xl_range(
		start[0],
		start[1],
		start[0],
		len(array)+start[1]-1
		)
	).Value = array

# Função para gravar planilha
def printdf(sheet,start, df):
	data = np.array(df)
	data = np.insert(data,0,list( df.columns ),axis=0)
	printxl(sheet,start, data)

# Função para abrir planilha
def openSheet(wb,sheet):
	try:
		return  wb.Worksheets(sheet)
	except Exception as e:
	 	ws = wb.Worksheets.Add()
	 	ws.Name = sheet
	 	return ws

# Abrir arquivo e planilha
wb = openWorkbook(os.getcwd()+'\exemplo.xlsx')
planilha =  wb.Worksheets("Plan1")
# Selecionar dados entre linhas 1 e 4 e colunas A e C e imprimir na tela
dados_lidos = dataFramexl(planilha, "A1:C4")
print(dados_lidos)

# Escrever texto na planilha começando na célula F1
lista = ['coluna1','coluna2']
printxl_line(planilha,'F1',lista)
# Escrever tabela em uma planilha - se não existir, é criada uma nova planilha
planilha_nova = openSheet(wb,'Plan2')
printdf(planilha_nova,'A1',dados_lidos)
