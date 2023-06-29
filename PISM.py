from __future__ import print_function
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
from time import sleep
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os.path
import string

options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")
options.add_experimental_option("excludeSwitches", ['enable-automation'])

navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)

#Define authorization scope / Definir escopo da autorização
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

#Define spreadsheet by its ID and cell intervall / Definir planilha pelo seu ID e intervalo de células
SAMPLE_SPREADSHEET_ID = ''
SAMPLE_RANGE_NAME = 'Name!A:R'

#Access authorization token / Acessar token de autorização 
creds = None
if os.path.exists('token.json'):
  creds = Credentials.from_authorized_user_file('token.json', SCOPES)
service = build('sheets', 'v4', credentials=creds)

#Access spreasheet and get values / Acessar planilha e extrair valores 
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()    

#Auxiliary variables / Variáveis auxiliares
cont = 1
notas = []
linhas = []

#List with uppercase alphabet letters / Lista com letras maiúsculas do alfabeto 
alphabet = list(string.ascii_uppercase)

#For each letter, get the information / Para cada letra, obter as informações
for j in range(len(alphabet)):

  print("Inicial: ", alphabet[j])
  navegador.get("http://www4.vestibular.ufjf.br/2022/notaspism1_aposrevisao/" + alphabet[j] +".html")
  #The number of students is defined by the number of details-control elements / O número de estudantes é definido pelo número de elementos details-control
  alunos = navegador.find_elements(By.CLASS_NAME, "details-control")
  sleep(5)

  aux = 1

  #If the script crashes for some reason, this control can be used to get the information starting from a specific point on the page
  #Se o programa para por alguma razão, esse controle pode ser usado para obter a informação começando de um ponto específico na página
  # if (j == Letra):
  #   aux = Ponto inicial
  # else:
  #   aux=1

  #For the information to be visible, the user must click on every '+' button / Para a informação ser disponibilizada, o usuário deve ser clicar em cada um dos botões '+'
  for i in range(1, len(alunos)):
    alunos[i].click()
    print("Clique", i)

  #With all the information now visible on the page, get it / Com toda as informações agora disponíveis no site, obtê-las
  for i in range(aux, len(alunos)):

    print(i)
    #Student name
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str( 2*i -1 ) + ']/td[2]').text)
    
    #Multiple-choice questions
    #Portuguese / Língua portuguesa
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[5]/td[3]').text)
    #Math / Matemática
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[6]/td[3]').text)
    #Literature / Literatura
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[7]/td[3]').text)
    #History / História
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[8]/td[3]').text)
    #Geography / Geografia
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[9]/td[3]').text)
    #Physics / Física
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[10]/td[3]').text)
    #Quimics / Química
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[11]/td[3]').text)
    #Biology / Biologia
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[12]/td[3]').text)
    
    #Open-ended questions
    #Portuguese / Portugês
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[13]/td[3]').text)
    #Math / Matemática
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[14]/td[3]').text)
    #Literature / Literatura
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[15]/td[3]').text)
    #History / História
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[16]/td[3]').text)
    #Geography / Geografia
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[17]/td[3]').text)
    #Physics / Física
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[18]/td[3]').text)
    #Quimics / Química
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[19]/td[3]').text)
    #Biology / Biologia
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[20]/td[3]').text)
    #Total
    linhas.append(navegador.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[' + str(2*i) +  ']/td/table/tbody/tr[21]/td/b[3]').text)

    notas.append(linhas)
 
    #Write the results / Escrever os resultados
    result = sheet.values().update(spreadsheetId = SAMPLE_SPREADSHEET_ID, range ='NAME!A' + str(cont + 1) + ':R' + str((cont + 1)), valueInputOption = 'USER_ENTERED', body = {'values': notas}).execute()

    linhas = []
    notas = []
    cont = cont + 1