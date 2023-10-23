# IMPORTANTE: ESTA AUTOMAÇÃO TRABALHA APENAS COM DADOS FICTICIOS QUE DEVERÃO SER SUBSTITUIDOS A PARTIR DA ABA E COLUNAS DA PLANILHA AO QUAO O USUÁRIO QUEIRA UTILIZAR.
# ALÉM DE SER NECESSÁRIO ALTERAR AS COORDENADAS DOS CLIQUES DE ACORDO COM A RESOLUÇÃO DO MONITOR.

import openpyxl
import pyperclip
import pyautogui
import time

work = openpyxl.load_workbook('produtos_ficticios.xlsx')
work_sheet = work['Produtos']
for linha in work_sheet.iter_rows(min_row=2):
    #nome do produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(2260,352, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #descrição do produto
    descricao_produto = linha[1].value
    pyperclip.copy(descricao_produto)
    pyautogui.click(2250,433, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #categoria do protudo
    categoria_produto = linha[2].value
    pyperclip.copy(categoria_produto)
    pyautogui.click(2282,580, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #código do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(2252,657, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #peso do produto
    peso_produto = linha[4].value
    pyperclip.copy(peso_produto)
    pyautogui.click(2268,750, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #dimensões do produto
    dimensoes_produto = linha[5].value
    pyperclip.copy(dimensoes_produto)
    pyautogui.click(2286,825, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(2264,903)
    time.sleep(3)

    #preço do produto
    preco_produto = linha[6].value
    pyperclip.copy(preco_produto)
    pyautogui.click(2260,380, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #quantidade de produtos
    quantidade_produto = linha[7].value
    pyperclip.copy(quantidade_produto)
    pyautogui.click(2250,465, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #validade do produto
    validade_produto = linha[8].value
    pyperclip.copy(validade_produto)
    pyautogui.click(2264, 560, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #cor do produto
    cor_produto = linha[9].value
    pyperclip.copy(cor_produto)
    pyautogui.click(2252,638, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #tamanho do produto
    tamanho_produto = linha[10].value
    pyautogui.click(2292,719,duration=1)
    if tamanho_produto == 'Pequeno':
        pyautogui.click(2286, 762, duration=1)
    elif tamanho_produto == 'Médio':
        pyautogui.click(2286, 790, duration=1)
    else:
        pyautogui.click(2286, 815, duration=1)

    #material do produto
    material_produto = linha[11].value
    pyperclip.copy(material_produto)
    pyautogui.click(2286, 811, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(2277,868)
    time.sleep(3)

    #fabriicante do produto
    fabricante_produto = linha[12].value
    pyperclip.copy(fabricante_produto)
    pyautogui.click(2255,402, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #país do produto
    pais_produto = linha[13].value
    pyperclip.copy(pais_produto)
    pyautogui.click(2260,487, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #observações do produto
    observacoes_produto = linha[14].value
    pyperclip.copy(observacoes_produto)
    pyautogui.click(2260, 568, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #código de barras do produto
    codigo_barras_produto = linha[15].value
    pyperclip.copy(codigo_barras_produto)
    pyautogui.click(2260, 706, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    armazem_produto = linha[16].value
    pyperclip.copy(armazem_produto)
    pyautogui.click(2260, 800, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(2277, 868)
    pyautogui.click(3098,619)
    pyautogui.click(2873,627)
