import sys
from design.login_design import *
from PyQt5.QtWidgets import  QMainWindow , QApplication, QFileDialog
import pyperclip  # ctrl+c e ctrl+v
import pyautogui  # mouse + keyboard
import os
from time import sleep
import webbrowser

class Sistema(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)
        self.btn_linkedin.clicked.connect(self.abri_linkedin)

    def abri_linkedin(self):
        get_url=webbrowser.open('https://www.linkedin.com/in/gabriel-jalles-ferreira-mota-9b7bb7200/')

    def abri_grh(self, login, senha):
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

if __name__=='__main__':
    qt = QApplication(sys.argv)
    sistema = Sistema()
    sistema.show()
    qt.exec_()
