import win32com.client as win32
import requests

from tkinter import *


# Criar a integracao com o Outlook
outlook = win32.Dispatch ('outlook.application')

# Criar um e-mail
email = outlook.CreateItem(0)

# Configurar as informações do seu e-mail
email.To = "juanbrasilcelular@gmail.m"
email.Subject = ('assunto')
email.HTMLBody = """
<p> Prezados, boa tarde! </p>

<p>Atenciosamente, </p>
"""

email.Send()
print ("E-mail enviado")

janela = Tk()
janela.title("Enviar Documentos Automaticamente")

texto_orientacao = Label(janela, text="Enviar Documento por Correio Eletrônico")
texto_orientacao.grid(column=0, row=0, padx=10, pady=10)

# Campo Assunto e Entrada de dados do Assunto
assunto = Label(janela, text="Digite o assunto do e-mail: ")
assunto.grid(column=0, row=1, padx=10, pady=10)

self.entrada = janela.Entry(self.janela)
self.entrada.pack(side=janela.LEFT, padx=10, pady=10)
self.entrada.bind("<Return>", self.info_assunto)

info_assunto (self, event)
    
#dados_assunto = janela.Entry(assunto)
#dados_assunto.pack(side=janela.LEFT, padx=10, pady=10)
#dados_assunto.bind("<Return>",)

botao = Button(janela, text="Enviar", command = email.Send)
botao.grid(column=0, row=3, padx=10, pady=10)



#enviar_para = Label(janela, text="Para:")
#enviar_para.grid(column=0, row=1)

janela.mainloop()