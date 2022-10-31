from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import datetime
import urllib # lib para formatação de variáveis (ajuste de tipagem)
import os
import numpy as np  # lib para formatação de variáveis (ajuste de tipagem)
import pandas as pd
import win32com.client as win32 # lib de envio de emails
from tkinter import *
from tkinter.ttk import *
from contacts import Contact
from sendEmail import sendEmail
from sendingFunctions import sendArquivo, sendImagem, sendEmail, sendTextMessage
from lblQuestion import MinhaFame

# Realiza a leitura da tabela
tabela = pd.read_excel('Envios.xlsx')

# função que converte o telefone em string
def to_str(var):
    return str(list(np.reshape(np.asarray(var), (1, np.size(var)))[0]))[1:-1]

def getCurrentDate():
    currentDate = datetime.datetime.now()
    textCurrentDate = str(currentDate.day) + '/' + str(currentDate.month)
    return textCurrentDate

def preview():
    # lista contatos
    printContacts()

def getFilteredContacts():
    validContacts = []

    for linha in tabela.index:
        nomeFromExcel = tabela.loc[linha, 'nome']
        mensagemFromExcel = tabela.loc[linha, 'mensagem']
        arquivoFromExcel = tabela.loc[linha, 'arquivo']
        imagemFromExcel = tabela.loc[linha, 'imagem']
        telefoneFromExcel = tabela.loc[linha, 'telefone']
        # converte o telefone em string
        telStr = to_str(telefoneFromExcel)

        paiFromExcel = tabela.loc[linha, 'pai']
        maeFromExcel = tabela.loc[linha, 'mae']
        generoFromExcel = tabela.loc[linha, 'genero']
        grupoFromExcel = tabela.loc[linha, 'grupo_especifico']
        birthdayFromExcel = tabela.loc[linha, 'nascimento']


        filterDict = handleFilter(varPai, varMae, varGenero, varGrupo)
        paiFilter = filterDict['pai']
        maeFilter = filterDict['mae']
        generoFilter = filterDict['genero']
        grupoFilter = filterDict['grupo']
        birthdayFilter = filterDict['birthday'] # Irrelevante, Sim ou Não

        textCurrentDate = getCurrentDate()

        if ((paiFromExcel == paiFilter or paiFilter == 'Irrelevante')
                and (maeFromExcel == maeFilter or maeFilter == 'Irrelevante')
                and (generoFromExcel == generoFilter or generoFilter == 'Irrelevante')
                and (grupoFromExcel == grupoFilter or grupoFilter == 'Irrelevante')
                and (birthdayFilter == "Sim" and birthdayFromExcel == textCurrentDate or birthdayFilter == 'Irrelevante')
                and len(telStr) == 13):
            validContact = Contact(nomeFromExcel, mensagemFromExcel, arquivoFromExcel, imagemFromExcel, telefoneFromExcel, birthdayFromExcel)
            validContacts.append(validContact)

    return validContacts

def printContacts():
    janelaFilteredData = Toplevel()
    # resolução da janela
    largura = 900
    altura = 600

    # resolução do sistema
    screen_width = janelaFilteredData.winfo_screenwidth()
    screen_height = janelaFilteredData.winfo_screenheight()

    # posição da janela
    posx = screen_width / 2 - largura / 2
    posy = screen_height / 2 - altura / 2
    janelaFilteredData.geometry('%dx%d+%d+%d' % (largura, altura, posx, posy))
    janelaFilteredData.title('Contatos Filtrados')

    Label(janelaFilteredData, text='Contatos Filtrados', width=65,font="Arial, 20", foreground='black', anchor=CENTER)\
        .grid(row=0, column=0, columnspan=5)

    columns = ('colNome', 'colMensagem', 'colArquivo', 'colImagem', 'colTelefone')
    tree = Treeview(janelaFilteredData, columns=columns, show='headings')

    tree.heading('colNome', text='Nome')
    tree.heading('colMensagem', text='Mensagem')
    tree.heading('colArquivo', text='Arquivo')
    tree.heading('colImagem', text='Imagem')
    tree.heading('colTelefone', text='Telefone')
    tree.column('colNome', width=150, anchor=CENTER)
    tree.column('colMensagem', width=300, anchor=CENTER)
    tree.column('colArquivo', width=150, anchor=CENTER)
    tree.column('colImagem', width=150, anchor=CENTER)
    tree.column('colTelefone', width=150, anchor=CENTER)


    tree.grid(row=1, column=0)
    btnConfirm = Button(janelaFilteredData, text='Confirmar envio das mensagens', command=lambda: sendMessages(janelaFilteredData))
    btnConfirm.grid(row=2, column=0, padx=10, pady=10, columnspan=5)

    btnCancel = Button(janelaFilteredData, text='Cancelar Filtragem', command=janelaFilteredData.destroy)
    btnCancel.grid(row=3, column=0, padx=10, pady=10, columnspan=5)

    for item in tree.get_children():
        tree.delete(item)
    validContacts = getFilteredContacts()
    for contato in validContacts:
        tree.insert('', 'end',
                    values=(contato.nome, contato.mensagem, contato.arquivo, contato.imagem, contato.telefone))

def sendMessages(janelaFilteredData):
    janelaFilteredData.destroy()
    navegador = webdriver.Chrome()
    navegador.get('https://web.whatsapp.com/')

    while len(navegador.find_elements(By.ID, 'side')) < 1:  # mantém testando enquanto não encontrat o ID = 'side'
        time.sleep(1)
    time.sleep(2)  # só uma garantia

    # armazena a lista de contatos filtrados na variável listOfValidContacts
    validContacts = getFilteredContacts()

    # itera a lista de contatos já filtrados e envia as mensagens

    for linha in validContacts:
        nome = linha.nome
        mensagem = linha.mensagem
        telefone = linha.telefone
        arquivo = linha.arquivo
        imagem = linha.imagem

        # enviar mensagem de texto
        sendTextMessage(nome, mensagem, telefone, navegador)

        # anexar arquivo
        sendArquivo(arquivo, navegador)

        # anexar imagem
        sendImagem(imagem, navegador)

    time.sleep(5)

    # envia o email com número de contatos que foram enviadas mensagens
    sendEmail(validContacts)

OPTIONS_GENERO = [
    "Irrelevante", "Irrelevante",
    "Masculino",
    "Feminino",
]

OPTIONS_MAE = [
    "Irrelevante", "Irrelevante", "Sim", "Não",
]

OPTIONS_PAI = [
    "Irrelevante", "Irrelevante", "Sim", "Não",
]

OPTIONS_GRUPO = [
    "Irrelevante", "Irrelevante", "Amigo", "Família", "Colega_trabalho", "Colega_infinity"
]

OPTIONS_BIRTH = [
    "Irrelevante", "Irrelevante", "Sim"
]

def getPai(varPai):
    userChoicePai = varPai.get()
    userLblPai.config(text=userChoicePai)
    return userChoicePai

def getMae(varMae):
    userChoiceMae = varMae.get()
    userLblMae.config(text=userChoiceMae)
    return userChoiceMae

def getGenero(varGenero):
    userChoiceGenero = varGenero.get()
    userLblGenero.config(text=userChoiceGenero)
    return userChoiceGenero

def getGrupo(varGrupo):
    userChoiceGrupo = varGrupo.get()
    userLblGrupo.config(text=userChoiceGrupo)
    return userChoiceGrupo

def getBirth(varBirth):
    userChoiceBirth = varBirth.get()
    userLblBirth.config(text=userChoiceBirth)
    return userChoiceBirth

def handleFilter(varPai, varMae, varGenero, varGrupo):

    pai = getPai(varPai)
    mae = getMae(varMae)
    genero = getGenero(varGenero)
    grupo = getGrupo(varGrupo)
    birthday = getBirth(varBirth)
    filterDict = {'pai': pai, 'mae': mae, 'genero': genero, 'grupo': grupo, 'birthday': birthday}
    return filterDict

# cria janela onde serão inseridos dados pelo usuário e mostrados contatos filtrados

janela = Tk()

# resolução da janela
largura = 1200
altura = 600

# resolução do sistema
screen_width = janela.winfo_screenwidth()
screen_height = janela.winfo_screenheight()

# posição da janela
posx = screen_width/2 - largura/2
posy = screen_height/2 - altura/2

janela.geometry('%dx%d+%d+%d' % (largura, altura, posx, posy))
janela.title('Sistema de Envio de Mensagens automáticas')

lblTitle = Label(janela,
                 width=35,
                 text="Sistema de Envio de Mensagens",
                 font="Arial 30",
                 foreground="black",
                 justify='center',
                 anchor='center'
                 )
lblTitle.grid(row=0, column=0, columnspan=3, padx=0, pady=0)

style = Style()

style.configure('TButton', font=
                    ('calibri', 20, 'bold'),
                    borderwidth='4')
style.map('TButton', foreground=[('active', '!disabled', 'green')],
                    background =[('active', 'black')])


style.configure('TLabel', background='white', width=50, justify="left", foreground='blue', font=('Arial', 12))

lblTitleQuestions = Label(janela, text="Filtros de Envio",
                          foreground='black', width=50, borderwidth=2, relief=SOLID, background='white', anchor=CENTER)
lblTitleQuestions.grid(row=1, column=0)

lblTitleOpções = Label(janela, text='Opção',
                       foreground='black', width=20, borderwidth=2, relief=SOLID, background='white', anchor=CENTER)
lblTitleOpções.grid(row=1, column=1)

lblTitleUserChoice = Label(janela, text='Escolha Usuário',
                           foreground='black', width=20, borderwidth=2, relief=SOLID, background='white', anchor=CENTER)
lblTitleUserChoice.grid(row=1, column=2)

# validação pai

lblPai = Label(janela, text='O contato tem que ser pai (Sim, Não, Irrelevante)? ', style='TLabel')
lblPai.grid(row=2, column=0)

varPai = StringVar()

varPai.set(OPTIONS_PAI[0])
vp = OptionMenu(janela, varPai, *OPTIONS_PAI)
vp.grid(row=2, column=1)

# btnPai = Button(janela, text="Confirmar", command=lambda: getPai(varPai))
# btnPai.grid(row=0, column=2)

userLblPai = Label(janela, textvariable=varPai, width=20, background='white', anchor=CENTER)
userLblPai.grid(row=2, column=2)

# validação mãe
lblMae = Label(janela, text='O contato tem que ser mãe (Sim, Não, Irrelevante)?', style='TLabel')
lblMae.grid(row=3, column=0)

varMae = StringVar()
varMae.set(OPTIONS_MAE[0])

vm = OptionMenu(janela, varMae, *OPTIONS_MAE)
vm.grid(row=3, column=1)

# btnMae = Button(janela, text="Confirmar", command=lambda: getMae(varMae))
# btnMae.grid(row=1, column=2)

userLblMae = Label(janela, textvariable=varMae, width=20, background='white', anchor=CENTER)
userLblMae.grid(row=3, column=2)

# validação gênero
lblGenero = Label(janela, text='É do sexo Masculino ou Feminino? ', style='TLabel')
lblGenero.grid(row=4, column=0)

varGenero = StringVar()
varGenero.set(OPTIONS_GENERO[0])

vg = OptionMenu(janela, varGenero, *OPTIONS_GENERO)
vg.grid(row=4, column=1)

# btnGenero = Button(janela, text="Confirmar", command=lambda: getGenero(varGenero))
# btnGenero.grid(row=2, column=2)

userLblGenero = Label(janela, textvariable=varGenero, width=20, background='white', anchor=CENTER)
userLblGenero.grid(row=4, column=2)

# validação grupo_especifico
lblGrupo = Label(janela, text='Pertence a algum grupo específico? ', style='TLabel')
lblGrupo.grid(row=5, column=0)

varGrupo = StringVar()
varGrupo.set(OPTIONS_GRUPO[0])

vgr = OptionMenu(janela, varGrupo, *OPTIONS_GRUPO)
vgr.grid(row=5, column=1)

userLblGrupo = Label(janela, textvariable=varGrupo, width=20, background='white', anchor=CENTER)
userLblGrupo.grid(row=5, column=2)

# data de aniversário
lblBirth = Label(janela, text='Deseja filtrar pelos aniversariantes de hoje? ', style='TLabel')
lblBirth.grid(row=6, column=0)

varBirth = StringVar()
varBirth.set(OPTIONS_BIRTH[0])

vbirth = OptionMenu(janela, varBirth, *OPTIONS_BIRTH)
vbirth.grid(row=6, column=1)

userLblBirth = Label(janela, textvariable=varBirth, width=20, background='white', anchor=CENTER)
userLblBirth.grid(row=6, column=2)

# btn Filtrar
btnConfirmFilter = Button(janela, text='Confirmar Filtragem', command=lambda: handleFilter(varPai, varMae, varGenero, varGrupo))

# btn Previsão
btnPreview = Button(janela, text='Ver Previsão', command=preview)
btnPreview.grid(row=7, column=0, padx=10, pady=10, columnspan=3)

janela.mainloop()