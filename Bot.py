#Importando as bibliotecas necessárias;
#Biblioteca para trabalhar planilhas.
import xlrd
from selenium import webdriver
#from whatsapp_api import WhatsApp
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import copy
import numpy as np
import os

WP_LINK = 'https://web.whatsapp.com'

## XPATHS
CONTACTS = '//*[@id="main"]/header/div[2]/div[2]/span'
SEND = '//*[@id="main"]/footer/div[1]/div[3]'
MESSAGE_BOX = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
NEW_CHAT = '//*[@id="side"]/header/div[2]/div/span/div[2]/div'
FIRST_CONTACT = '//*[@id="app"]/div/div/div[2]/div[1]/span/div/span/div/div[2]/div/div/div/div[2]/div'
SEARCH_CONTACT ='//*[@id="app"]/div/div/div[2]/div[1]/span/div/span/div/div[1]/div/label/div/div[2]'
#This is my components in contribution;
ATTACHMENT = '//*[@id="main"]/header/div[3]/div[1]/div[2]/div[1]'
ATTACHMENT_IMAGE = '//*[@id="main"]/header/div[3]/div[1]/div[2]/span/div[1]/div[1]/ul/li[1]/button'
ATTACHMENT_IMAGE_URL = '//*[@id="main"]/header/div[3]/div[1]/div[2]/span/div[1]/div[1]/ul/li[1]/button/input'
RETURN = '//*[@id="app"]/div[1]/div[1]/div[2]/div[1]/span/div[1]/header/div[1]/div[1]/button'

class WhatsApp:
    # O local de execução do nosso script
    dir_path = os.getcwd()
    # O caminho do chromedriver
    chromedriver = os.path.join(dir_path, "chromedriver.exe")
    # Caminho onde será criada pasta profile
    profile = os.path.join(dir_path, "profile", "wpp")
    def __init__(self):
        self.driver = self._setup_driver()
        self.driver.get(WP_LINK)
        print("Please scan the QR Code")

    @staticmethod
    def _setup_driver():
        print('Loading...')
        options = Options()
        options.add_argument("disable-infobars")
        driver = webdriver.Chrome(options=options)
        return driver

    def _get_element(self, xpath, attempts=5, _count=0):
        '''Safe get_element method with multiple attempts'''
        try:
            element = self.driver.find_element_by_xpath(xpath)
            #print('Found element!')
            return element
        except Exception as e:
            if _count<attempts:
                sleep(1)
                print(f'Attempt {_count}')
                self._get_element(xpath, attempts=attempts, _count=_count+1)
            else:
                print("Cheguei até o comando de clicar no botão para voltar")
                try:
                    print("Tentando voltar")
                    self._click(RETURN)
                    pass
                except Exception as e:
                    print("Não achei o botão de retornar")    
                print("Element not found")
                pass

    def _click(self, xpath):
        el = self._get_element(xpath)
        try:
            el.click()
        except NoSuchElementException:
            print("Não encontrei o elemento")
            pass

    def _send_keys(self, xpath, message):
        el = self._get_element(xpath)
        el.send_keys(message)

    def send_message(self, message):
        '''Write and send message'''
        self.write_message(message)
        self._click(SEND)

#Essa funcionalidade lê, corrigi e grava nomes e números;
def ReadExcel():
    #Eu buscarei esses valores dentro da string para removê-lo, afim de evitar erros quando realizar a pesquisa do telefone celular no whattsap;
    FirstInvalidValues = "+"
    SecondInvalidValues = "."
    #Variável que armazenará o nome em análise atual;
    ActualNumber = ""
    #Variável que armazenará o nome em análise atual;
    ActualNickname = ""
    QtdDeLinhas = 0

    #Farei um primeiro laço para percorrer o excel formatando os números e removendo o que houver de formatações incorretos;
    for rx in range(sh.nrows):
        #Atribuirei o número de telefone a uma variável para facilitar o seu tratamento e manuseio, e transformarei em string para poder compará-lo e alterá-lo se necessário;
        #Sempre farei a leitura da coluna b(1), e não considero que exista um cabeçalho, pois sempre começo a consumir pela linha 0, como se ali já houvesse o primeiro valor;
        ActualNumber = str(sh.cell_value(rx, 1))

        #Adiciona um novo elemento ao array;
        nicknames.append(str(sh.cell_value(rx, 2)))    

        pos=0
        for pos in range(0, len(ActualNumber)):
            #Se encontrar o símbolo da adição, o algoritmo assumirá que encontrou um +55, portanto removerá os três elementos da string;
            if ActualNumber[pos] == FirstInvalidValues:
                ActualNumber = ActualNumber[4:(len(ActualNumber)-1)]
                break

        pos=0
        for pos in range(0, len(ActualNumber)):
            #Se encontrar um símbolo de ponto, o algoritmo assumirá que encontrou um número com a formatação de float, e portanto removerá os dois últimos dígitos assumindo que são .0;
            if ActualNumber[pos] == SecondInvalidValues:
                ActualNumber = ActualNumber[0:len(ActualNumber)-2]
                break

        numbers.append(ActualNumber)

    contact = 0
    while(contact < len(numbers)):
        
        contact += 1
    print("A planilha foi lida com sucesso!")
    return numbers, nicknames    


def WhatsappDealing( attempts=5, _count=0):
    #Inicializando um objeto do whatsapp;
    wp = WhatsApp()
    sleep(4)
    contact = 0
    notRegistered = 0
    while(contact < len(numbers)):
        print(contact)
        print(nicknames[contact])
        #print(contact)
        #print(notRegistered)
        sleep(2)
        #butt = wp.driver.find_element_by_xpath(NEW_CHAT)
        #butt.click()
        try:
            wp._click(NEW_CHAT)
            sleep(2)
        except Exception as e:
            el = wp.driver.find_element_by_css_selector('div._2EXPL')
            el.click()
            sleep(5)
            #wp.send_files()
            sleep(3)
            try:
                # Clica no botão adicionar
                wp.driver.find_element_by_css_selector("span[data-icon='clip']").click()
                # Seleciona input
                attach = wp.driver.find_element_by_css_selector("input[type='file']")
                imagem = wp.dir_path + "/us.jpeg"
                # Adiciona arquivo
                attach.send_keys(imagem)
                sleep(3)
                # Seleciona botão enviar
                send = wp.driver.find_element_by_xpath("//div[contains(@class, 'yavlE')]")
                # Clica no botão enviar
                send.click()
            except Exception as e:
                print("Erro ao enviar media", e)
            sleep(3)
            wp.send_message(theMessage)         
            continue   
        wp._send_keys(SEARCH_CONTACT, numbers[contact])
        sleep(2)
        try:
            el = wp.driver.find_element_by_xpath(FIRST_CONTACT)
            el.click()
            sleep(5)

            time.sleep(5)
            sleep(3)
            try:
                # Clica no botão adicionar
                wp.driver.find_element_by_css_selector("span[data-icon='clip']").click()
                # Seleciona input
                attach = wp.driver.find_element_by_css_selector("input[type='file']")
                imagem = wp.dir_path + "/us.jpeg"
                # Adiciona arquivo
                attach.send_keys(imagem)
                sleep(3)
                # Seleciona botão enviar
                send = wp.driver.find_element_by_xpath("//div[contains(@class, 'yavlE')]")
                # Clica no botão enviar
                send.click()
            except Exception as e:
                print("Erro ao enviar media", e)
            sleep(3)
            wp.send_message(theMessage) 
            #el = wp.driver.find_element_by_class("_2EXPL _1f1zm")
            sleep(2)
        except Exception as e:
            butt = wp.driver.find_element_by_css_selector('button._1aTxu')
            butt.click()
            sleep(3)
            if contact < len(numbers):
                numbersNotRegistered.append(str(numbers[contact]))
                namesNotRegistered.append(str(nicknames[contact]))
                numbers.pop(contact)
                nicknames.pop(contact)
                notRegistered +=1
                contact += 1 
                continue
            else:
                break
        if _count<attempts:
            sleep(1)
        else:
            print("Cheguei até o comando de clicar no botão para voltar")
        time.sleep(1)
        contact += 1   
    imprime()


def imprime():
    print(nicknames)
    print(numbers)
    print(namesNotRegistered)
    print(numbersNotRegistered)

def getTheMessage():
    print("Olá, insira abaixo a mensagem que você deseja enviar aos seus diversos contatos:")
    print("Instruções: \n 1 - Escreva a mensagem toda aqui, nós buscaremos os contatos telefônicos ou nomes dos contatos (da maneira que estiverem salvo em seus contatos) na coluna B; \n 2 - Na coluna C devem estar os nomes das pessoas que serão substituídos em sua frase; \n 3 - A planilha não deve conter cabeçalho (título dos dados); \n 4 - Na mensagem insira parenteses (), para que esses símbolos sejam substituídos pelos nomes da planilha;")
    mensagem = str(input())
    print()
    print("Reultado: ")
    print(mensagem)
    print("Confirma? (S/N)")
    confirma = str(input())
    if confirma == "S":
        mensagem.replace("()", "'+ str(nicknames[contact]) +'")
        theMessage = mensagem
        print("A mensagem foi interpretada e armazenada com sucesso!")
    elif confirma == "N":
            getTheMessage()
    else:
        print("Não identificamos sua resposta, por favor insira sua mensagem novamente!")
        getTheMessage()
    
    


#Se um dia em quiser que essa solução seja escalável, devo criar aqui uma lógica para abrir a janela e o usuário selecionar onde o arquivo de contatos está;
book = xlrd.open_workbook("contacts.xls")
#Carregando a primeira guia do excel. Caso construa uma solução escalável, isso precisa ser flexibilizado;
sh = book.sh = book.sheet_by_index(0)
#Array que receberá todos os números de telefone da lista já tratados;
numbers = []
#Array que receberá todos os nomes da lista;
names = []
#Array que receberá todos os apelidos da base;
nicknames = []
#Números não cadastrados;
numbersNotRegistered = []
#Nomes não cadastrados;
namesNotRegistered = []
#A mensagem;
theMessage = ""

resultSaves = ""
resultNotSaves = ""

def SalvarResultado():
    resultNotSaves = "Os contatos não salvos são: \n"
    #str(numbersNotRegistered).strip('[]')
    i=0
    while i < len(numbersNotRegistered):
        resultNotSaves = resultNotSaves + "nome" + str(namesNotRegistered[i]) + ", número " + str(numbersNotRegistered[i]) +"\n"
        i += 1

    arquivo = open('contacts.txt', 'w')
    arquivo.write(resultNotSaves)
    arquivo.close()

#ReadExcel()
getTheMessage()
#WhatsappDealing()       
#SalvarResultado()