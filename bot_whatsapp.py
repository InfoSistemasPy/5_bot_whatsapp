#!/usr/bin/env python
# coding: utf-8

# In[1]:


#===================Bibliotecas==================

import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui
import time
import random
import urllib
import os

#==================Funções===================
def Caminho_App():
    texto = os.getcwd()
    texto = texto.replace('\\', '/') 
    caminho_completo = texto + "/img"
    print(caminho_completo)
    time.sleep(2)
    
def Entrar_Selecionar_Ok():
    pyautogui.click(x=1452, y=59)
    time.sleep(3)
    pyautogui.write(caminho_completo)
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.doubleClick(x=290, y=219)
    time.sleep(3)
    button = nave.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span')
    button.click()


#========================Loops==========================    
def Enviar_Mensagens():
    for i, mensagem in enumerate(agenda['Mensagem']):
        pessoa = agenda.loc[i,"Pessoa"]
        numero = agenda.loc[i, "Número"]
        texto = urllib.parse.quote(f"Olá {pessoa}! {mensagem}")
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        nave.get(link)
        while len(nave.find_elements_by_id("side")) < 1:
            time.sleep(5)
        nave.find_element_by_xpath('/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        time.sleep(5)
        nave.find_element_by_xpath('/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
        time.sleep(5)
        nave.find_element_by_xpath('/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button/span').click()
        time.sleep(5)
        Entrar_Selecionar_Ok()   
        time.sleep(30)

texto = os.getcwd()
texto = texto.replace('\\', '/') 
caminho_completo = texto + "/conteudo"
print(caminho_completo)
time.sleep(2)
    
#===============WebDriver_Selenium================

nave = webdriver.Edge()

nave.get('https://web.whatsapp.com/')
time.sleep(30)

while len(nave.find_elements_by_id("side")) < 1:
    time.sleep(1)

#============Importando Agenda.xls===========

agenda = pd.read_excel("Agenda.xlsx")
display(agenda)

#==================Chamar Funções==================

Enviar_Mensagens()
    

#import time
#import pyautogui
#time.sleep(4)
#pyautogui.position()