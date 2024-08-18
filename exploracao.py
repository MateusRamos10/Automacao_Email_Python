from imap_tools import MailBox, AND
from pandas import DataFrame as df
import pandas as pd
from io import BytesIO
from datetime import datetime

# https://github.com/ikvk/imap_tools/blob/master/README.rst#guide
# https://docs.python.org/3/library/smtplib.html
# https://docs.python.org/3/library/email.mime.html#email.mime.base.MIMEBase
# https://www.systoolsgroup.com/imap/


# imap-mail.outlook.com
# usuario = "mateusramos.oficial@gmail.com"
# senha = "cejs pbde cfqn uwnt"

# Class Encontrar Email

# ----------------------- def Loga no Email
usuario = input("Digite seu email: ")
senha = input("Digite sua senha: ")
meu_email = MailBox("imap.gmail.com").login(usuario, senha)

destinatario = input("Digite o Email do Destinatáio: ")
assunto = input("Digite o assunto para filtrar: ")

# ----------------------- def lista emails

# lista_emails = meu_email.fetch(AND(from_='brunacamsilva@gmail.com', subject='Python'))
lista_emails = meu_email.fetch(AND(from_=destinatario, subject=assunto))

lista_cabecalho = []

for email in lista_emails:
    titulo_email = f"{email.subject}"
    data = f"{email.date}"

    if email.attachments:
        for anexo in email.attachments:
            anexo_titulo = f"{anexo.filename}"
            anexo_conteudo = anexo.payload
            
            lista_cabecalho.append((titulo_email, data, anexo_titulo, anexo_conteudo))
    else:
        lista_cabecalho.append((titulo_email, data, "Sem anexo", "Sem Conteudo"))

# ----------------------- def cria dataframe
title = ['Titulo','Data','Anexo','Conteúdo']

resultado = df(lista_cabecalho, columns=title, index=None)
resultado = resultado.astype({'Titulo': 'object', 'Data': 'datetime64', 'Anexo': 'string'})
resultado = resultado.sort_values(by='Data',ascending=False)
resultado['Data'] = resultado['Data'].dt.strftime('%d-%m-%Y')

resultado.iloc[:, :-1]

# ----------------------- def escolha email
email_escolhido = int(input("Digite o código do email escolhido: "))

if 0 <= email_escolhido < len(lista_cabecalho):
    print(email_escolhido)
    anexo_escolhido = lista_cabecalho[email_escolhido][3]

# ----------------------- def edita anexo
dados_em_bytes = anexo_escolhido
buffer = BytesIO(dados_em_bytes)
planilha_anexo = pd.read_excel(buffer)
planilha_anexo_modificado = planilha_anexo.drop(['ASDAD','ASDASD','ASDASD.1','ASDSAD'], axis=1)
planilha_anexo_modificado
data_atual = datetime.now()
data_atual = data_atual.strftime('%d-%m-%Y')
arquivo_nome = f'arquivo{data_atual}.xlsx'

# ----------------------- def cria documento
caminho_arquivo = f'C:\\GitHub\Automacao_Email_Python\\{arquivo_nome}'  # Substitua com o caminho real
planilha_anexo_modificado.to_excel(caminho_arquivo, index=False)

# ----------------------- Class Enviar Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


destinos = input("Digite o Email que será enviado: ")

# ----------------------- def criacao_email()
def criacao_email():
    # Configurações do e-mail
    corpo_email = """
    <p>Email enviado automaticamente usando <b>Python</b></p>
    <p>Contém <b>Anexo</b></p>
    """
    
    # Criar mensagem multipart
    msg = MIMEMultipart()
    msg['Subject'] = 'Email Automático usando Python'
    msg['From'] = usuario
    msg['To'] = destinos
    # msg['To'] = 'brunacamsilva@gmail.com'

    # Corpo do e-mail
    msg.attach(MIMEText(corpo_email, 'html'))
    
# ----------------------- def carregar_anexo    
    with open(caminho_arquivo, 'rb') as anexo_arquivo:
        anexo = MIMEApplication(anexo_arquivo.read(), _subtype='xlsx')
        anexo.add_header('Content-Disposition', 'attachment', filename=arquivo_nome)
        msg.attach(anexo)

# ----------------------- def enviar_email
    with smtplib.SMTP('smtp.gmail.com', 587) as s:
        s.starttls()
        s.login(usuario, senha)
        s.send_message(msg)
    
    print('Email Enviado')

# ---------------------------------------------------------------------------------------------------
criacao_email()
