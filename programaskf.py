#coding: UTF-8

from platform import python_branch
from selenium import webdriver                              #Importa a biblioteca Selenium
from webdriver_manager.chrome import ChromeDriverManager    #Importa a biblioteca que gerencia automaticamente a versão do chrome e selenium sem precisar instalar
from selenium.webdriver.chrome.service import Service       #Importa Service 
from selenium.webdriver import Chrome                       #Importa o Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
import pyautogui                                            #Importa a biblioteca de automação
import datetime                                             #Importa a biblioteca data atual
import time                                                 #Importa a biblioteca time



def OpenWebSite():
    global navegador                                                            #Declara navegador como uma variável global (Que pode ser utilizada em outras funções)
    servico = Service(ChromeDriverManager().install())                          
    navegador = webdriver.Chrome(service=servico)                               #Abre o navegador
    navegador.get("https://repcenter.skf.com/machineviewer/logon.aspx")         #Entra no site SKF
    
    

        
    '''


    #DESATIVADO
    
    #ABRIR O SITE COM PYAUTOGUI
    
    
    
    pyautogui.moveTo(35,5)                                  #Move para a barra de tarefas
    pyautogui.click()                                       #Clica
    pyautogui.write("goog", interval=0.25)                  #Digita na barra de tarefas goog
    time.sleep(1)
    pyautogui.press("enter")                                #Pressiona ENTER
    time.sleep(10)
    pyautogui.moveTo(400,50)            
    pyautogui.click()                                       #Clica
    time.sleep(1)
    pyautogui.write("https://repcenter.skf.com/machineviewer/logon.aspx")     #Digita o site
    pyautogui.press("enter")            #Pressiona ENTER
    time.sleep(10)
    '''


def Login():
    
    login_path = '#identif_UserName'                                                    #VARIÁVEL PARA GUARDAR A CÉLULA DO LOGIN
    password_path = '#identif_Password'                                                 #VARIÁVEL PARA GUARDAR A CÉLULA DA SENHA

    login_element = navegador.find_element_by_css_selector(login_path)                  #ACESSA A CÉLULA DO LOGIN
    password_element = navegador.find_element_by_css_selector(password_path)            #ACESSA A CÉLULA DA SENHA

    login_element.send_keys('thomaz.silva')                                             #DIGITA O LOGIN
    password_element.send_keys('CXvm.r/$A[$Q>65E')                                      #DIGITA A SENHA
    
    navegador.find_element_by_xpath('//*[@id="identif_LoginButton"]').click()           #CLICA PARA LOGAR
    
    navegador.implicitly_wait(25)                                                                     #ESPERA 25 SEGUNDOS

    
 

    
    
    
    '''
    #FAZENDO LOGIN COM PYAUTOGUI (ANTIGO)
    pyautogui.moveTo(1000,500)                              #Move para o local do login
    pyautogui.click()                                       #Clica
    pyautogui.write("thomaz.silva",interval=0.10)           #Digita o login com intervalo
    time.sleep(2)
    pyautogui.press("tab")                                  #Muda para digitar a senha
    time.sleep(2)
    pyautogui.write("CXvm.r/$A[$Q>65E",interval=0.10)       #Escreve a senha
    pyautogui.press("enter")                                #Pressiona ENTER  
    time.sleep(20)
    '''

def GetDate():
    global hoje                                                                                         #TORNA HOJE UMA VARIÁVEL GLOBAL
    hoje = datetime.date.today()                                                                        #VARIÁVEL QUE HOJE (CONTÉM A DATA ATUAL)
    print(hoje)

def TreatDate():
    global treated_date
    
    treated_date = hoje.strftime('%d/%m/%Y')                                              #TRATA A DATA COMO DEVE SER ESCRITA 
    return treated_date


def Get_Worksheet():                                                                                 
    
    navegador.find_element_by_xpath('//*[@id="Reports"]/span[1]').click()    #CLICA NA CÉLULA RELATÓRIOS
    navegador.implicitly_wait(30) 
    
     
    navegador.find_element_by_css_selector("button[class ='MuiButtonBase-root MuiButton-root MuiButton-outlined MuiButton-outlinedPrimary']").click()               #FECHA A CAIXA DE SPAM 
    
    
    
    navegador.implicitly_wait(30)                                               # É UM TIME.SLEEP QUE SE ENCONTRAR O ELEMENTO EXECUTA ANTES
    navegador.find_element_by_xpath('//*[@id="detailedAssetHealth"]').click()                           #CLICA NA CÉLULA SAÚDE DETALHADA DO ATIVO
    navegador.implicitly_wait(30)  
    navegador.find_element_by_xpath('//*[@id="panel1a-content"]/div[1]/form/div[2]/div[1]/div/select').click()  #SELECIONA A HIERARQUIA
    navegador.implicitly_wait(30)  
    navegador.find_element_by_xpath('//*[@id="panel1a-content"]/div[1]/form/div[2]/div[1]/div/select/option[2]').click()  #SELECIONA BUNGE RIO GRANDE
    navegador.implicitly_wait(30)  
    

    initial_date_path = '#startDate'                                                      #VARIÁVEL PARA GUARDAR A CÉLULA DE INSERIR DATA INICIAL 
    #final_date_path = '#endDate'                                                         #VARIÁVEL PARA GUARDAR A CÉLULA DE INSERIR DAT'A FINAL
    
    initial_date_element = navegador.find_element_by_css_selector(initial_date_path)
    #final_date_element = navegador.find_element_by_css_selector(final_date_path)
                                                         
    #Não usado 
    # final_date_element = navegador.find_elements_by_css_selector(final_date_path)           #ACESSA A CÉLULA DATA FINAL


    navegador.implicitly_wait(30)
    initial_date_element.send_keys(Keys.CONTROL + "a")                                        #DA UM CONTROL A PARA SELECIONAR TODA A LINHA             
    initial_date_element.send_keys(Keys.DELETE)                                               #DA UM DELETE E LIMPA A LINHA
    initial_date_element.send_keys("29/01/2020")                                              #ESCREVER A DATA INICIAL  

    #ERRO
    navegador.find_element_by_xpath('//*[@id="panel1a-content"]/div[1]/form/div[4]/div[3]/div/div[2]/div/label/span[2]').click()   #DESMARCA A CAIXA CONCLUÍDOS
    #ERRO
    time.sleep(100)
    navegador.find_element_by_xpath('//*[@id="panel1a-content"]/div[2]/div/button/span[1]').click()         #CLICA NO LISTA DETALHADA DE ATIVOS

    

    #final_date_element.clear()
    #final_date_element.send_keys(hoje)                                                        #ESCREVER A DATA FINAL
    
    navegador.implicitly_wait(30)





    
    navegador.maximize_window()
    pyautogui.moveTo(1120,350)
    pyautogui.click()
    pyautogui.scroll(+10000)                                                                  #SCROLLA A PÁGINA PRA CIMA
    pyautogui.moveTo(1120,430)
    pyautogui.click()
    time.sleep(2)
    
    
   
 



    '''
    #EXTRAIR COM PYAUTOGUI(ANTIGO)
    pyautogui.moveTo(100,350)                               #Move para a aba relatórios
    pyautogui.click()                                       #Clica
    time.sleep(1)
    pyautogui.moveTo(480,180)                               #Seleciona a aba Saúde detalhada do ativo
    pyautogui.click()                                       #Clica
    time.sleep(1)
    
    pyautogui.moveTo(500,450)                               #Seleciona a caixa
    pyautogui.click()                                       #Clica
    time.sleep(1)
    pyautogui.moveTo(500,500)                               #Seleciona a Bunge
    pyautogui.click()                                       #Clica
    
    pyautogui.moveTo(1340,680)                              #Seleciona a última caixa
    pyautogui.click()                                       #Clica
    '''
    



def ExtractInformation():
    GetDate()
    TreatDate()
    OpenWebSite()
    Login()
    Get_Worksheet()
    

   

    time.sleep(100)
    
    



ExtractInformation()