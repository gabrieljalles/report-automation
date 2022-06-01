from time import sleep
import pyautogui
import os
from operator import itemgetter
import grh
import pyperclip
import shutil


def limpa_arq_xlsx(nome_relatorio):
    arquivo_temp = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
    os.remove(arquivo_temp)

def existe_data(data_final):
    nome_arquivo = open(r'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Hora extra\confere.txt', 'r')
    achou = ''
    for x in nome_arquivo:
        x = x.strip()
        if x == data_final:
            achou= '1'
            break
        else:
            achou= '0'
    nome_arquivo.close()
    return achou

def escreve_arquivo(data_final):
    nome_arquivo = open(r'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Hora extra\confere.txt', 'a')
    nome_arquivo.write(f'{data_final}\n')
    nome_arquivo.close()

def hora_extra(data_final):
    achou = existe_data(data_final)
    if achou == '0':
        verba = '0180'
        grh.tirar_verba(data_final, verba)
        nome_relatorio = f'hora extra'
        grh.baixar_relatorio(nome_relatorio)
        pyautogui.hotkey('alt', 'f4')
        origem = fr'C:\..arelatoriotemporario\{nome_relatorio}.xlsx'
        origem_original = r'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Hora extra\Verba 0180 hora extra.xlsx'
        sleep(0.5)
        os.startfile(origem)
        sleep(0.5)
        sleep(0.5)
        pyautogui.press('down')
        sleep(2)
        pyautogui.hotkey('ctrl', 't')
        pyautogui.hotkey('ctrl', 'c')
        sleep(3)
        os.startfile(origem_original)
        sleep(8)
        excel_habilitar_edicao = pyautogui.locateCenterOnScreen('imagens_automacao/excel_habilitar_edicao.png', confidence=0.5)
        pyautogui.moveTo(excel_habilitar_edicao)
        pyautogui.click(excel_habilitar_edicao)
        sleep(5)
        pyautogui.hotkey('ctrl', 'up')
        sleep(2)
        pyautogui.hotkey('ctrl', 'down')
        sleep(2)
        pyautogui.press('down')
        sleep(2)
        pyautogui.hotkey('ctrl', 'v')
        sleep(2)
        pyautogui.press('down')
        pyautogui.press('up')
        pyautogui.hotkey('ctrl', '-')
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'b')
        sleep(15)
        pyautogui.hotkey('alt', 'f4')
        sleep(5)
        pyautogui.hotkey('alt', 'f4')
        sleep(2)
        limpa_arq_xlsx(nome_relatorio)
        escreve_arquivo(data_final)
    else:
        print('A atualização já foi feita para esse mês')


    # destino = fr'\\172.22.1.26\Users\DA\Ger Adm\Sup TI CDOC\N TI\powerbi\GRH\BI 2.0\Verbas\Adicional Periculosidade\{nome_relatorio}.xlsx'
    # shutil.move(origem, destino)
    # sleep(4)
    # grh.limpa_arq(nome_relatorio)


hora_extra('022022')
