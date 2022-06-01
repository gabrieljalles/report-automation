from time import sleep
import pyautogui
import os
from operator import itemgetter
import grh
import pyperclip
import shutil


def falta(data_final):
    verba = '0520'
    grh.tirar_verba(data_final, verba)
    data_final_formatada = []
    for x in data_final:
        data_final_formatada += ''.join(x)
    if data_final_formatada.insert(2, "-") is None:
        data_final_formatada = ''.join(itemgetter(0, 1, 2, 3, 4, 5, 6)(data_final_formatada))
        nome_relatorio = f'Faltas {data_final_formatada}'
        grh.baixar_relatorio(nome_relatorio)
        sleep(3)
        alteracao_tabela_verba_0520 = pyautogui.locateCenterOnScreen('imagens_automacao/alteracao_tabela_verba_0520.png', confidence=0.8)
        pyautogui.doubleClick(alteracao_tabela_verba_0520)
        pyperclip.copy(r'Verba 520 (Faltas)01-2017;12-20')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        sleep(0.5)
        pyautogui.hotkey('alt', 'f4')
        sleep(0.5)
        pyautogui.press('enter')
        origem = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
        destino = fr'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Faltas\{nome_relatorio}.xlsx'
        shutil.move(origem, destino)
        sleep(4)
        grh.limpa_arq(nome_relatorio)


falta('032022')
