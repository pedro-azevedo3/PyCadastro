# Esse programa em Python irá ler dados de uma planilha
# Inserir cada célula de cada linha em um campo do sistema
# Lembrando que para rodar você precisa de acesso ao aplicativo de cadastro do sistema.
import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('venda_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

# Coordenadas pegas usando Mouse Info!

for linha in vendas_sheet.iter_rows(min_row=2):
    # Nome
    pyautogui.click(1808,452,duration=1.5)
    pyautogui.write(linha[0].value)
    # Produto
    pyautogui.click(1815,476,duration=1.5)
    pyautogui.write(linha[1].value)
    # Quantidade
    pyautogui.click(1813,497,duration=1.5)
    pyautogui.write(str(linha[2].value))
    # Categoria
    pyautogui.click(1883,532,duration=1.5)
    pyautogui.write(linha[3].value)
    # Salvar
    pyautogui.click(1752,549,duration=1.5)
    pyautogui.write(linha[3].value)
    # OK 
    pyautogui.click(1256,581,duration=1.5)
    pyautogui.write(linha[3].value)  
    