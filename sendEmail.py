import win32com.client as win32 # lib de envio de emails

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