import time
from selenium.webdriver.common.by import By
import win32com.client as win32 # lib de envio de emails
import os
import urllib # lib para formatação de variáveis (ajuste de tipagem)

def sendImagem(imagem, navegador):
    if imagem != "N":
        #   clica no clips de anexar
        caminho_completo = os.path.abspath(f"{imagem}")
        navegador.find_element(By.XPATH,
                               '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()

        # preenche o input do anexar documento
        navegador.find_element(By.XPATH,
                               '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input').send_keys(
            caminho_completo)
        time.sleep(2)

        # clica na seta de confirmar envio
        navegador.find_element(By.XPATH,
                               '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span').click()

def sendArquivo(arquivo, navegador):
    if arquivo != "N":
        #   clica no clips de anexar
        caminho_completo = os.path.abspath(f"{arquivo}")
        navegador.find_element(By.XPATH,
                               '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
        # preenche o input do anexar documento
        navegador.find_element(By.XPATH,
                               '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input').send_keys(
            caminho_completo)
        time.sleep(2)
        # clica na seta de confirmar envio
        navegador.find_element(By.XPATH,
                               '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span').click()

def sendEmail(validContacts):
    nrValidContacts = len(validContacts)
    print(nrValidContacts)

    # criar a intergração com o OutLook
    outlook = win32.Dispatch('outlook.application')

    # criar um e-mail
    email = outlook.CreateItem(0)

    # configurar as informações do seu emial
    email.To = "tencopello@hotmail.com"
    email.Subject = "e-mail automático do WhatsApp Auto"
    email.HTMLBody = f"""
    <body>
        <h3>Olá <b>Leonardo</b>, aqui é o sistema de envio de mensagens WhatsApp</h3>

        <p>Foram enviadas mensagens para <b>{nrValidContacts}</b> contatos</p>
        <p>Os contatos foram:</p>

        <h3>Att.,</h3>

        <h3>Sistema de Envios Automático</h3>
    </body>
    """
    email.Send()
    print('Email enviado')

    def sendTextMessage(nome, mensagem, telefone, navegador):
        texto = mensagem.replace('fulano', nome)  # substitui a palavra fulano pelo nome dele
        texto = urllib.parse.quote(texto)  # ??????? Transforma o texto do excel em formato aceito pelo whatsApp
        link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
        #     navegador = webdriver.Chrome()

        navegador.get(link)  # vai para tela de envio de mensagens com os dados do contato já inseridos no whatsAppWeb

        # esperar a tela de whatsApp carregar -> espear um elemento que só existe na tela carregada aparecer
        while len(navegador.find_elements(By.ID, 'side')) < 1:  # mantém testando enquanto não encontrar o ID = 'side'
            time.sleep(1)
        time.sleep(2)  # só uma garantia

        # enviar mensagem
        navegador.find_element(By.XPATH,
                               '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        time.sleep(3)  # coloquei provisoriamente

def sendTextMessage(nome, mensagem, telefone, navegador):
    texto = mensagem.replace('fulano', nome)  # substitui a palavra fulano pelo nome dele
    texto = urllib.parse.quote(texto)  # ??????? Transforma o texto do excel em formato aceito pelo whatsApp
    link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
    #     navegador = webdriver.Chrome()

    navegador.get(link)  # vai para tela de envio de mensagens com os dados do contato já inseridos no whatsAppWeb

    # esperar a tela de whatsApp carregar -> espear um elemento que só existe na tela carregada aparecer
    while len(navegador.find_elements(By.ID, 'side')) < 1:  # mantém testando enquanto não encontrar o ID = 'side'
        time.sleep(1)
    time.sleep(2)  # só uma garantia

    # enviar mensagem
    navegador.find_element(By.XPATH,
                           '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
    time.sleep(3)  # coloquei provisoriamente