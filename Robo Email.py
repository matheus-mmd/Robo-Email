import PySimpleGUI as sg
from imap_tools import MailBox, AND, A
import time



class Robo_email:

    def __init__(self):
        layout = [
            [sg.Text('*Usuario:', size=(15, 1)), sg.Input(key='Usuario', size=(35, 1))],
            [sg.Text('*Senha:  ',size=(15, 1)), sg.Input(key='Senha',password_char='*', size=(35, 1))],
            [sg.Text('*Imap:    ',size=(15, 1)), sg.Input(key='Imap', size=(35, 1))],
            [sg.Text('*Assunto do Email:',size=(15, 1)), sg.Input(key='Assunto', size=(35, 1))],
            [sg.Text('Texto PDF: ',size=(15, 1)), sg.Input(key='PDF', size=(35, 1))],
            [sg.Text('*Tempo para Loop:',size=(15, 1)),sg.Input(key='Tempo', size=(35, 1))],
            [sg.Button('Logar',size=(15, 1))]
            
        ]
        self.layout_login = sg.Window('Robo Email', layout=layout)

    def janela(self):
        while True:
            evento, valores = self.layout_login.read()
            if evento == sg.WINDOW_CLOSED:
                break

            if evento == 'Logar':
                usuario = valores['Usuario']
                senha = valores['Senha']
                imap = valores['Imap']
                assunto = valores['Assunto']
                pdf = valores['PDF']
                tempo = int(valores['Tempo'])

                if usuario != '' and senha != '' and imap != '' and assunto != '' and tempo != '':
                    while True:
                        with MailBox(imap).login(usuario, senha) as mailbox:
                            for msg in mailbox.fetch(AND(subject=assunto)):
                                print(msg.date, len(msg.text or msg.html))
                                if len(msg .attachments) > 0:
                                    for anexo in msg.attachments:
                                        if pdf != '':
                                            if pdf in anexo.filename:
                                                with open(anexo.filename, 'wb') as arquivo_excel:
                                                    arquivo_excel.write(anexo.payload)
                                        else:
                                            with open(anexo.filename, 'wb') as arquivo_excel:
                                                arquivo_excel.write(anexo.payload)

                            time.sleep(tempo)


re = Robo_email()
re.janela()

# imap.mailcorp.com.br
