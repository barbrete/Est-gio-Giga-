import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
import pyautogui as pag
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import win32com.client
import os

# Ainda falta colocar ele pra rodar pra sempre e queria fazer uma verificação para
# ver se os sites estão fora do ar ou nao

service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
urlAIR = 'urlAIR'
urlOFS = 'urlOFS'
email = 'email'
password = 'password'

def aumenta_tela():
    for _ in range(3):
        pag.hotkey('ctrl','+')
        
def diminui_tela():
    for _ in range(5):
        pag.hotkey('ctrl','-')
        
def arruma_tela():
    pag.hotkey('ctrl','0')

driver.implicitly_wait(15)
#queria colocar um if else ou do while para verificar se está conectado na conta desses dois sistemas e se
#nao estivesse ia entrar nesse loop e conectar
driver.get(urlAIR)
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[1]/section/form/div/div/input').send_keys(email)
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[1]/section/form/button').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[1]/div/form/div/div/input').send_keys(password)
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[1]/div/form/button').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div/div[2]/div/div[2]/div[4]/div[2]').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[1]/a').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[2]/div[5]/section/div[3]/input').click() #R8
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[2]/div[5]/section/div[6]/input').click() #R9
time.sleep(2)

footer = driver.find_element(By.TAG_NAME, "footer")
delta_y = footer.rect['y']
ActionChains(driver)\
    .scroll_by_amount(0, 200)\
    .perform()

driver.find_element(By.XPATH, '/html/body/div/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[2]/div[7]/section/div[1]/input').send_keys('C')
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[2]/div[7]/section/div[10]/input').click()

driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[2]/input').click()
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[2]/input').click()
time.sleep(2)
campo_texto = driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[1]/input')
campo_texto.send_keys('ativa')
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[3]/input').click() #ativacao
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[6]/input').click() #ativacao pap
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[7]/input').click() #ativacao empresarial
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[8]/input').click() #ativacao pme
time.sleep(2)
campo_texto.clear()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[1]/input').send_keys('clean')
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[3]/input').click() #clean up
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[4]/input').click() #clean up mudanca de endereco
time.sleep(2)
campo_texto.clear()
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[1]/input').send_keys('reparo')
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[3]/input').click() #reparo
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[4]/input').click() #reparo empresarial
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[9]/input').click() #reparo pme
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[10]/input').click() #reparo preventivo
time.sleep(1)
campo_texto.clear()
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[1]/input').send_keys('upgr')
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[3]/input').click() #upgrade
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[5]/input').click() #upgrade nao logico
time.sleep(1)
campo_texto.clear()
campo_texto.send_keys('mudan')
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[3]/input').click() #mudanca de endereco
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/div/div/div[1]/div[5]/section/div[4]/input').click() #mudanca de comodo
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[3]/div[1]/div[2]/div/div/div[3]/div/div/form/footer/button[1]').click() #filtrar
time.sleep(1)

def salvar_excel(dados): 
    if dados:
        df = pd.DataFrame(dados)
        df.to_excel("dados_extraidos.xlsx", index=False, engine='openpyxl')
        print("Dados salvos em 'dados_extraidos.xlsx'!")
    else:
        print("Nenhum dado foi coletado.")
        
def pegar_informacao():
    driver.find_element(By.XPATH, '/html/body//div[2]/div/div/nav/div[2]/div/a[2]/span[2]').click()
    numero_contrato = driver.find_element(By.XPATH, '/html/body//div[2]/div/div/div[2]/div[2]/div/div//div[2]/p[3]/span/span').text
    pag.press('ESC')

    return numero_contrato
    
def verificar_conveniencia() :
    quadradinho_para_formatar = '//*[@id="app"]/div/section/div[3]/div[2]/div[2]/table/tbody/tr[{numero}]/td[19]/span/span/button'
    aba = '//*[@id="app"]/div/section/div[3]/div[2]/div[3]/div[2]/div/div/ul/li[{numero}]'
    pagina = True 
    j = 1
    dados = []
    
    while (pagina):
        try:         
            for i in range(15):
                xpath_conveniencia = quadradinho_para_formatar.format(numero=i+1)
                time.sleep(1)
                
                try:
                    driver.find_element(By.XPATH, xpath_conveniencia).click()
                    time.sleep(2)                             
                    numero = pegar_informacao()

                    if numero is not None:
                        dados.append({"Número": numero})
                    else:
                        break  # Sai do loop quando não encontrar mais tópicos
                    
                except Exception as e:
                    print(f"Erro ao clicar no item {i+1}: {e}")
                    continue            
                
                
                
                # ActionChains(driver) \
                #     .move_by_offset(1450, 0) \
                #     .perform()
                # time.sleep(1)
             
            time.sleep(2)
        except Exception as e:
            print(f"Ocorreu um erro ao processar a página {j}: {e}")
            pagina = False
            
        j += 1 
        
        ActionChains(driver)\
            .scroll_by_amount(0, 200)\
            .perform()
        try:
            xpath_pagina = aba.format(numero=j)
            driver.find_element(By.XPATH, xpath_pagina).click()
            print(f"Indo para a página {j}")
        except Exception as e:
            print(f"Ocorreu um erro ao tentar acessar a página {j}: {e}")
            pagina = False  
                
    salvar_excel(dados)

diminui_tela()
verificar_conveniencia()
arruma_tela()
driver.quit()
