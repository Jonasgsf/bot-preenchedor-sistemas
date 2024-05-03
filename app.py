import openpyxl
import pyautogui

'''
comandos para abrir o mouse Info pelo terminal
python
from mouseinfo import mouseInfo
mouseInfo()
(salva a coordenada com f6)
'''
#ler o arquivo excel

planilha = openpyxl.load_workbook("bot_preenchedor-sistemas/vendas_de_produtos.xlsx")

# escolhendo a pagina específica

pagina_vendas = planilha['vendas']

#navegando no conteúdo da planilha

for row in pagina_vendas.iter_rows(min_row = 2, values_only=True):
    cliente, produto, quantidade, categoria = row
    #colocando as variaveis nos campos do sistema com o pyautogui
    #pelas coordenadas usando o mouseinfo
    #caso o esteja dando erro, converter o valor para string pode resolver

    pyautogui.click('x','y', duration=1.5)#duration é a duração da acão
    pyautogui.write(cliente)

    pyautogui.click('x','y', duration=1.5)#duration é a duração da acão
    pyautogui.write(produto)

    pyautogui.click('x','y', duration=1.5)#duration é a duração da acão
    pyautogui.write(str(quantidade))

    pyautogui.click('x','y', duration=1.5)#duration é a duração da acão
    pyautogui.write(categoria)

    #caso necessário, fazer etapas para 'click' em campos de salvar processo, etc.
    pyautogui.click('x','y', duration=1.5)#duration é a duração da acão 
    