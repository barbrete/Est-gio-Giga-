import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
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
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div/div[2]/div/div[2]/div[5]/div[2]').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[2]/div[1]/div/aside/ul/li[1]/ul/li[1]/a').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div/a[2]').click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[12]/td[4]/span').click()

time.sleep(2)

driver.execute_script("window.open('');")
abas = driver.window_handles
driver.switch_to.window(abas[1])
driver.get(urlOFS)
driver.find_element(By.XPATH, '//*[@id="sign-in-with-sso"]/div').click()
driver.find_element(By.XPATH, '//*[@id="sso_username"]').send_keys(email)
driver.find_element(By.XPATH, '//*[@id="continue-with-sso"]/div').click()
time.sleep(3)

if (driver.current_url.startswith('https://login.microsoftonline.com/')):
    try:
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="i0116"]').send_keys(email)
        driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="i0118"]').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="KmsiCheckboxField"]').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
    
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    finally:
        
        try:
            pag.press('F5')
        except Exception as e:
            print(f"Erro ao {e}")
    #fazer verificação para caso nao de certo de se conectar tentar novamente

time.sleep(10)
driver.find_element(By.XPATH, '//button[@aria-label="Exibir"]').send_keys(Keys.ENTER)
driver.find_element(By.XPATH, '//*[@id="ui-id-52"]').click()

start_element = driver.find_element(By.XPATH, '//*[@id="ui-id-52"]')

#Usando ActionChains para pressionar Tab 5 vezes e depois Enter
actions = ActionChains(driver)

# Mover o foco para o elemento inicial, caso necessário
actions.move_to_element(start_element).perform()

#Simular pressionamento de Tab 5 vezes
for _ in range(5):
    actions.send_keys(Keys.TAB).perform()

actions.send_keys(Keys.ENTER).perform()

time.sleep(5)

driver.find_element(By.XPATH, '//button[@aria-label="Ações"]').send_keys(Keys.ENTER)
driver.find_element(By.XPATH, '/html/body/div[26]/div/div/button[2]').send_keys(Keys.ENTER)

def verificar_se_existe():
    pasta = 'C:/Users/barbara.gianvechio/Downloads/'
    primeiro_arquivo = 'chamados_abertos_field_service.xlsx'
    segundo_arquivo = 'Atividades-Casa Cliente'

    if os.path.exists(os.path.join(pasta, primeiro_arquivo)):
        if any(arquivo.startswith(segundo_arquivo) for arquivo in os.listdir(pasta)):
            existe = True
        else:
            existe = False
    else:
        existe = False

    return existe
    
tem = verificar_se_existe()
demorou_quanto_tempo = 0
while (tem == False):
    time.sleep(60)    
    demorou_quanto_tempo = demorou_quanto_tempo + 1
    tem = verificar_se_existe()

driver.quit()
print(demorou_quanto_tempo)

def run_excel_macro():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(r'C:\Users\barbara.gianvechio\Desktop\BACKLOG - REGIONAL I E II.xlsm')

        excel.Application.Run("Backlog")

        time.sleep(5) 

        workbook.Close(SaveChanges=False)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    finally:
        try:
            excel.Quit()
        except Exception as e:
            print(f"Erro ao tentar fechar o Excel: {e}")

run_excel_macro()

nome_arquivo_OFS = 'Atividades-Casa Cliente'
nome_arquivo_AIR = 'chamados_abertos_field_service.xlsx'
pasta = r'C:/Users/barbara.gianvechio/Downloads/'
time.sleep(20)
def excluir_arquivos(AIR, OFS, pasta):
    for arquivo in os.listdir(pasta):
        caminho = os.path.join(pasta, arquivo)
        if (arquivo.startswith(AIR)):
            os.remove(caminho)
        elif (arquivo.startswith(OFS)):
            os.remove(caminho)  
                      
excluir_arquivos(nome_arquivo_AIR, nome_arquivo_OFS,pasta)

pag.PAUSE = 2.5

# tirar isso
pag.alert('Favor não usar o mouse ou o teclado enquanto o código executa.')

pag.press('winleft')
time.sleep(1)

pag.write('whatsapp')
time.sleep(2)

pag.press('enter')
contatos = pd.read_excel("Contatos.xlsx")

for i, mensagem in enumerate(contatos['mensagem']):
    pessoa = contatos.loc[i,"nome"]    
    numero = contatos.loc[i,"numero"]
    pag.hotkey('ctrl','f') 
    pag.write(str(pessoa))
    pag.press('tab')
    pag.press('enter')
    pag.hotkey('winleft','e')
    pag.press('right')
    pag.press('enter')
    pag.press('down')
    pag.press('up')
    pag.hotkey('ctrl','c')
    pag.hotkey('alt','F4')
    pag.hotkey('ctrl','v')
    pag.press('enter')

pag.alert('Código finalizado.')
