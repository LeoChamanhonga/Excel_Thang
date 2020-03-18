def excel_file(self, df_final, filename, sheet_name):
	df = df_final[['data','valor']].iloc[1:6,]
	df.to_excel(filename, sheet_name=sheet_name, index=False)



