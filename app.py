import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(352,368,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(352,463,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(354,588,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(353,675,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(353,759,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(353,844,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(352,905,duration=1)
    sleep(1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(347,390,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(347,478,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(347,562,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(352,649,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(352,729,duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(352,769,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(352,793,duration=1)
    else:
        pyautogui.click(352,819,duration=1)
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(352,827,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(352,882,duration=1)
    sleep(1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(352,420,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(352,500,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(352,594,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(352,715,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(352,801,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(352,860,duration=1)
    pyautogui.click(1131,272,duration=1)
    sleep(1)

    pyautogui.click(954,637,duration=1)
    sleep(1)

#to open mouse info

#python
#from mouseinfo import mouseInfo
#mouseInfo()
#https://cadastro-produtos-devaprender.netlify.app/index.html