import pyautogui
import os
from time import sleep
import cv2
import grh
import shutil


def efetivo_comissionado(data_final):
    # Dados
    # ----------------------------------------
    data_inicial = '01011900'
    nome_relatorio = 'Efetivo-comissionado'
    btn_relatorio = pyautogui.locateCenterOnScreen('imagens_automacao/btn_relatorio.png', confidence=0.8)
    pyautogui.click(btn_relatorio)
    sleep(1)
    btn_relatorio_institucionais = pyautogui.locateCenterOnScreen('imagens_automacao/btn_relatorio_institucionais.png',confidence=0.8)
    pyautogui.moveTo(btn_relatorio_institucionais)
    sleep(1)
    btn_btn_relatorio_institucionais_relacaogeralsemagrupamento = pyautogui.locateCenterOnScreen('imagens_automacao/btn_relatorio_institucionais_relacaogeralsemagrupamento.png',)
    pyautogui.click(btn_btn_relatorio_institucionais_relacaogeralsemagrupamento)
    pyautogui.write(data_inicial)
    pyautogui.write(data_final)
    pyautogui.press('down')
    pyautogui.press('tab')
    pyautogui.write('00000000')
    pyautogui.write('99999999')
    pyautogui.press('tab')
    pyautogui.press('down')
    pyautogui.press('tab')
    pyautogui.press('tab')
    registro_recuperado = None
    while registro_recuperado is None:
        registro_recuperado = pyautogui.locateCenterOnScreen('imagens_automacao/relatorio_efetivo_comissionado_pronto.png',)
    else:
        grh.baixar_relatorio(nome_relatorio)
        pyautogui.hotkey('alt', 'f4')
        origem = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
        destino = fr'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\GRH  working\{nome_relatorio}.xlsx'
        shutil.move(origem, destino)
        grh.limpa_arq(nome_relatorio)


    #salvar no c:/pasta nova



if __name__ == "__main__":
    efetivo_comissionado('18052022')
