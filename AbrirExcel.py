#
from win32com.client import Dispatch

'''Para invocar a aplicação recorre-se ao objecto Dispatch, importado no exemplo anterior, tendo como argumento a aplicação que se pretende
manipular e recebe-se um handler que contém a ligação à aplicaçã '''
xlApp = Dispatch("Excel.Application")

'''Para podermos observar a janela, é necessário torná-la visível, para tal atribui-se o valor 1 à propriedade Visible.
Afectar o valor 0 à propriedade torna a janela invisível. '''
xlApp.Visible = 1


#Para terminar a aplicação recorre-se ao método Quit(), sendo assim fechada a aplicação.
xlApp.Quit()


# Por fim, liberta-se o handler que mantinha a ligação
del xlApp


'''O Excel Object Model indica quais os objectos, propriedades e métodos
que se encontram disponíveis para efectuar a manipulação de um documento Excel.
De forma simles, Application refere a aplicação, Workbook refere um documento,
Worksheet refere uma página,
Cell refere uma célula e Range refere um conjunto de células. '''



#Assim, para criar um novo documento,
#basta dizer à aplicação que esta deve adicionar um novo documento:
xlApp.Workbooks.Add()

'''Criar uma nova folha ao documento é bastante simples,
recorrendo ao documento activo,
referenciado por ActiveWorkbook, adiciona-se uma nova folha ao mesmo '''

xlSheet = xlApp.ActiveWorkbook.Worksheets.Add()

'''Atribuir ou ler valores às células de um folha é igualmente simples,
basta referir a célula pela sua
coordenada e atribuir-lhe o valor pretendido, ou requisitar o valor actual: '''

#Atribuir Valor
xlSheet.Cells(1,1).Value = 'Exemplo'
#Ler Valor
print xlSheet.Cells(1, 1).Value
