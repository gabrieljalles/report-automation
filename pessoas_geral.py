import grh
from time import sleep
import pyautogui
import os
from operator import itemgetter
import grh
import shutil


def pessoas_geral(data_inicial, data_final):
    data_inicial_formatada = []
    for x in data_inicial:
        data_inicial_formatada += ''.join(x)
    data_inicial_formatada.insert(4, "-")
    data_inicial_formatada = ''.join(itemgetter(2, 3, 4, 5, 6, 7, 8)(data_inicial_formatada))
    nome_relatorio = f'Pessoas (geral) {data_inicial_formatada}'
    print(nome_relatorio)
    btn_relatorio = pyautogui.locateCenterOnScreen('imagens_automacao/btn_relatorio.png', confidence=0.8)
    pyautogui.click(btn_relatorio)
    pyautogui.press(['down', 'down', 'right','down','enter'])
    pyautogui.write(data_inicial)
    pyautogui.write(data_final)
    pyautogui.press(['down', 'down', 'down', 'down', 'down'])
    pyautogui.press('tab')
    pyautogui.write('00000000')
    pyautogui.write('99999999')
    pyautogui.press('tab')
    sleep(7)
    grh.baixar_relatorio(nome_relatorio)
    alteracao_tabela_pessoas_geral = pyautogui.locateCenterOnScreen('imagens_automacao/alteracao_tabela_pessoas_geral.png', confidence=0.8)
    pyautogui.doubleClick(alteracao_tabela_pessoas_geral)
    pyautogui.write('Pessoas (geral)')
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'f4')
    pyautogui.press('enter')
    origem = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
    destino = fr'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\GRH  working\Pessoas (geral)\{nome_relatorio}.xlsx'
    shutil.move(origem, destino)
    sleep(5)
    grh.limpa_arq(nome_relatorio)





if __name__ == "__main__":
    pessoas_geral('01032022','31032022')


