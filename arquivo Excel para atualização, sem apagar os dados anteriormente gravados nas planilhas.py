# -*- coding: utf-8 -*-
# Script para gravar em planilha Excel concatenando conteúdo

import pandas as pd
import numpy as np

# Criar data frame de dados aleatórios
df = pd.DataFrame(np.random.randint(0,100,size=(20, 4)), columns=list('ABCD'))

# Criar objeto para leitura e selecionar planilha
excel_reader = pd.ExcelFile('exemplo.xlsx')
to_update = {"Plan1": df}
# Criar objeto para escrita
excel_writer = pd.ExcelWriter('exemplo.xlsx')

# Loop para o caso de querer trabalhar com mais de uma planilha
for sheet in excel_reader.sheet_names:
	sheet_df = excel_reader.parse(sheet)
	append_df = to_update.get(sheet)
	# Concatenar com o que já existia
	if append_df is not None:
		sheet_df = pd.concat([sheet_df, df]).drop_duplicates()
	# Gravar no arquivo
	sheet_df.to_excel(excel_writer, sheet, index=False)
# Salvar e fechar arquivo
excel_writer.save()
