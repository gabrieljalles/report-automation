from time import sleep
import pyautogui
import os
from operator import itemgetter
import grh
import pyperclip
import shutil


def funcao_confianca(data_final):
        nome_relatorio = f'Funcao de confianca'
        btn_relatorio = pyautogui.locateCenterOnScreen('imagens_automacao/btn_relatorio.png', confidence=0.8)
        pyautogui.click(btn_relatorio)
        pyautogui.press(['down', 'right', 'up', 'up', 'up'])
        pyautogui.press('enter')
        pyautogui.write(data_final)
        pyautogui.press(['down', 'down', 'tab'])
        pyautogui.write('0')
        pyautogui.press('tab')
        pyautogui.write('99')
        pyautogui.press('tab')
        sleep(27)
        grh.baixar_relatorio(nome_relatorio)
        pyautogui.hotkey('alt', 'f4')
        origem = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
        destino = fr'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Função de confiança\{nome_relatorio}.xlsx'
        shutil.move(origem, destino)
        sleep(4)
        grh.limpa_arq(nome_relatorio)

funcao_confianca('032022')
