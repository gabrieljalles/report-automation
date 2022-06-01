from time import sleep
import pyautogui
import os
from operator import itemgetter
import grh
import pyperclip
import shutil


def adicional_insalubridade40(data_final):
    verba = '0153'
    grh.tirar_verba(data_final, verba)
    data_final_formatada = []
    for x in data_final:
        data_final_formatada += ''.join(x)
    if data_final_formatada.insert(2, "-") is None:
        data_final_formatada = ''.join(itemgetter(0, 1, 2, 3, 4, 5, 6)(data_final_formatada))
        nome_relatorio = f'Verba153 (AD de Insalubridade40%){data_final_formatada}'
        grh.baixar_relatorio(nome_relatorio)
        pyautogui.hotkey('alt', 'f4')
        origem = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
        destino = fr'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Adicional Insalubridade 40%\{nome_relatorio}.xlsx'
        shutil.move(origem, destino)
        sleep(4)
        grh.limpa_arq(nome_relatorio)


adicional_insalubridade40('042022')