import pyperclip  # ctrl+c e ctrl+v
import pyautogui  # mouse + keyboard
import os
import shutil
from time import sleep
import efetivo_comissionado
# ------- O programa só funcionou com esses modulos instalados---------
# pip install opencv-python
# pip install Pillow --upgrade


def abri_grh(login, senha):
    print('Abrindo GRH...')
    caminho_grh = r'C:\Sistemas\grh80\executa\a_grh.exe'
    os.startfile(caminho_grh)
    sleep(1)
    pyautogui.write(login)
    pyautogui.press('tab')  # selecionando a empresa  dmae ...
    sleep(1)
    botao_selecao_grh = pyautogui.locateOnScreen('imagens_automacao/botao_selecao_grh.png', confidence=0.7)
    pyautogui.click(botao_selecao_grh)
    sleep(1)
    selecao_dmae = pyautogui.locateOnScreen('imagens_automacao/selecao_dmae.png', confidence=0.7)
    pyautogui.click(selecao_dmae)
    sleep(1)
    pyautogui.press('tab')
    sleep(1)
    pyautogui.write(senha)
    pyautogui.press('enter')
    sleep(1)
    grass = pyautogui.locateOnScreen('imagens_automacao/folha_de_pagamento.png', confidence=0.7)
    sleep(1)
    pyautogui.click(grass)
    print('verificando se precisa atualizar o GRH ...')
    atualiza_grh()


def atualiza_grh():
        print("Tentando atualizar o grh !")
        sleep(5)
        aviso_desatualizado = pyautogui.locateCenterOnScreen('imagens_automacao/ok_desatualizado.png', confidence=0.6)
        if aviso_desatualizado is None:
            print("Seu grh já está atualizado !")
        else:
            pyautogui.click(aviso_desatualizado)
            sleep(5)
            btn_ok_atualizado_1 = None
            btn_ok_atualizado_2 = None
            while  btn_ok_atualizado_1 and btn_ok_atualizado_2 is None:
                print("Atualizando GRH...")
                btn_ok_atualizado_1 = pyautogui.locateCenterOnScreen('imagens_automacao/botao_atualizado_com_sucesso_ok.png')
                btn_ok_atualizado_2 = pyautogui.locateCenterOnScreen('imagens_automacao/botao_atualizado_com_sucesso_ok_2.png')
                sleep(3)
            else:
                sleep(20)
                print(btn_ok_atualizado_1, btn_ok_atualizado_2)
                pyautogui.click(btn_ok_atualizado_1)
                sleep(1)
                pyautogui.click(btn_ok_atualizado_2)
                print("Seu grh foi atualizado com sucesso !")

                abri_grh(input_login, input_password)

def sair_grh():
    print('Fechando o GRH...')
    botao_x_fecha_click = pyautogui.locateOnScreen('imagens_automacao/x_botao_fecha_click.png', confidence=0.7)
    pyautogui.click(botao_x_fecha_click)
    sleep(1)
    botao_aviso_desatualizado = pyautogui.locateOnScreen('imagens_automacao/aviso_desatualizado_botao_ok.png', confidence=0.95)
    pyautogui.click(botao_aviso_desatualizado)
    print('GRH fechado com sucesso !')


def baixar_relatorio(nome_relatorio):
    global caminho_baixado
    caminho_baixado = rf'C:\..arelatoriotemporario\{nome_relatorio}.xls'
    pyautogui.press('f12')
    sleep(2)
    pyautogui.press('esc') #sair primeira caixa
    pyautogui.press(['tab','tab','tab','tab']) #diretorio
    pyautogui.press('down') # abrir opcoes
    sleep(0.5)
    pyautogui.press(['pageup','pageup','pageup','pageup','pageup','pageup','pageup','pageup',])
    sleep(0.5)
    pyautogui.press(['d','d','d'])
    pyautogui.press('enter')
    pyautogui.press(['tab','tab'])
    pyautogui.press('down')
    pyautogui.press('up')
    pyautogui.press('enter')
    pyautogui.press(['tab', 'tab'])

    pyautogui.write(nome_relatorio)
    pyautogui.press('tab')
    pyautogui.press('down')
    pyautogui.press(['up','up','up','up','up','up','up','up','up'])
    sleep(2)
    pyautogui.press('tab')
    pyautogui.press('enter')

    substituicao = None
    cont=0
    while substituicao is None:
        substituicao = pyautogui.locateCenterOnScreen('imagens_automacao/aviso_substituicao.png', confidence=0.7)
        cont += 1
        if cont == 10:
            break
    else:
        pyautogui.press('s')

    excel_habilitacao(nome_relatorio)


def excel_habilitacao(nome_relatorio):
    os.startfile(rf'C:\..arelatoriotemporario\{nome_relatorio}.xls')
    sleep(5)
    excel_habilitar_edicao = pyautogui.locateCenterOnScreen('imagens_automacao/excel_habilitar_edicao.png', confidence=0.9)
    pyautogui.click(excel_habilitar_edicao)
    pyautogui.hotkey('ctrl', 'b')
    sleep(2)
    pyautogui.press('enter')
    sleep(0.5)
    pyautogui.press('enter')
    sleep(0.5)
    pyautogui.press('enter')
    sleep(0.5)
    pyautogui.press('enter')
    sleep(0.5)
    pyautogui.press('enter')
    sleep(0.5)
    pyautogui.press('left')
    sleep(0.5)
    pyautogui.press('enter')

def tirar_verba(data_final, verba):
    btn_relatorio = pyautogui.locateCenterOnScreen('imagens_automacao/btn_relatorio.png', confidence=0.8)
    pyautogui.click(btn_relatorio)
    pyautogui.press(['down', 'down', 'down'])
    pyautogui.press(['right', 'down','down','down','down','down','down','down','down', 'enter'])
    pyautogui.write(data_final)
    pyautogui.write(data_final)
    pyautogui.write(verba)
    pyautogui.write('01')
    pyautogui.press(['down', 'down', 'down', 'down', 'down', 'down', 'down'])
    pyautogui.press(['tab', 'tab', 'tab', 'tab', 'tab', 'tab'])
    sleep(5)
    pyautogui.press('enter')

def limpa_arq(nome_relatorio):
        arquivo_temp = fr'C:\..arelatoriotemporario\{nome_relatorio}.xls'
        os.remove(arquivo_temp)

if __name__ == "__main__":
    input_login = 'XXXXXXXXXXXXX'
    input_password = 'XXXXXXXXXX'
    abri_grh(input_login,input_password)

