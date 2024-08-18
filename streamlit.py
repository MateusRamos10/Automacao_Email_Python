import streamlit as st
from imap_tools import MailBox, AND
from pandas import DataFrame as df
import pandas as pd
from io import BytesIO
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Classe Encontrar_Email
class Encontrar_Email:
    def __init__(self, usuario, senha, servidor_imap):
        self.usuario = usuario
        self.senha = senha
        self.meu_email = MailBox(servidor_imap).login(usuario, senha)

    def lista_emails(self, destinatario, assunto):
        lista_emails = self.meu_email.fetch(AND(from_=destinatario, subject=assunto))
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

        return lista_cabecalho

    def cria_dataframe(self, lista_cabecalho):
        title = ['Titulo', 'Data', 'Anexo', 'Conteúdo']

        resultado = df(lista_cabecalho, columns=title, index=None)
        resultado = resultado.astype({'Titulo': 'object', 'Data': 'datetime64', 'Anexo': 'string'})
        resultado = resultado.sort_values(by='Data', ascending=False)
        resultado['Data'] = resultado['Data'].dt.strftime('%d-%m-%Y')
        resultado = resultado.iloc[:, :-1]
        resultado.index.name = 'ID'
        resultado = resultado.head(10)

        return resultado

    def edita_anexo(self, anexo_escolhido):
        dados_em_bytes = anexo_escolhido
        buffer = BytesIO(dados_em_bytes)
        planilha_anexo = pd.read_excel(buffer)
        planilha_anexo_modificado = planilha_anexo.drop(['ASDAD', 'ASDASD', 'ASDASD.1', 'ASDSAD'], axis=1)
        return planilha_anexo_modificado

# Classe Criacao_Email
class Criacao_Email:
    def __init__(self, usuario, senha, caminho_arquivo, destinos, servico_email):
        self.usuario = usuario
        self.senha = senha
        self.caminho_arquivo = caminho_arquivo
        self.destinos = destinos
        self.servico_email = servico_email

    def criacao_email(self):
        corpo_email = """
        <p>Email enviado automaticamente usando <b>Python</b></p>
        <p>Contém <b>Anexo</b></p>
        """

        msg = MIMEMultipart()
        msg['Subject'] = 'Email Automático usando Python'
        msg['From'] = self.usuario
        msg['To'] = self.destinos

        msg.attach(MIMEText(corpo_email, 'html'))

        with open(self.caminho_arquivo, 'rb') as anexo_arquivo:
            anexo = MIMEApplication(anexo_arquivo.read(), _subtype='xlsx')
            anexo.add_header('Content-Disposition', 'attachment', filename=self.caminho_arquivo.split('\\')[-1])
            msg.attach(anexo)

        return msg

    def enviar_email(self, msg):   
        if self.servico_email == 'Gmail':
            smtp_server = 'smtp.gmail.com'
        elif self.servico_email == 'Outlook':
            smtp_server = 'smtp-mail.outlook.com'
        else:
                st.error("Serviço de e-mail não suportado.")
                return

        try:
            with smtplib.SMTP(smtp_server, 587) as s:
                s.starttls()
                s.login(self.usuario, self.senha)
                s.send_message(msg)
            st.success('Email Enviado')
        except Exception as e:
            st.error(f'Erro ao enviar email: {e}')

# Função principal do Streamlit
def main():
    st.title('Automatizador de Email com Anexo')

    st.subheader('Login no Email')

    servico_email = st.selectbox("Escolha o serviço de email:", ["Gmail", "Outlook"])
    if servico_email == "Gmail":
        servidor_imap = "imap.gmail.com"
    elif servico_email == "Outlook":
        servidor_imap = "imap-mail.outlook.com" 

    usuario = st.text_input("Digite seu email:")
    senha = st.text_input("Digite sua senha:", type='password')

    if st.button('Login'):
        if usuario and senha:
            email_obj = Encontrar_Email(usuario, senha, servidor_imap)
            st.session_state['email_obj'] = email_obj
            st.success("Login realizado com sucesso!")
        else:
            st.error("Preencha o email e a senha corretamente!")

    if 'email_obj' in st.session_state:
        st.subheader('Filtrar Emails')
        destinatario = st.text_input("Digite o Email do Destinatário:")
        assunto = st.text_input("Digite o assunto para filtrar:")

        if st.button('Buscar Emails'):
            if destinatario and assunto:
                lista_cabecalho = st.session_state['email_obj'].lista_emails(destinatario, assunto)
                resultado = st.session_state['email_obj'].cria_dataframe(lista_cabecalho)
                st.session_state['lista_cabecalho'] = lista_cabecalho
                st.dataframe(resultado)
            else:
                st.error("Preencha o destinatário e o assunto!")

        if 'lista_cabecalho' in st.session_state:
            email_escolhido = st.number_input("Digite o código do email escolhido:", min_value=0, max_value=len(st.session_state['lista_cabecalho'])-1, step=1)
            if st.button('Editar Anexo'):
                anexo_escolhido = st.session_state['lista_cabecalho'][email_escolhido][3]
                planilha_anexo_modificado = st.session_state['email_obj'].edita_anexo(anexo_escolhido)
                data_atual = datetime.now().strftime('%d-%m-%Y')
                arquivo_nome = f'arquivo{data_atual}.xlsx'
                caminho_arquivo = f'C:\\GitHub\\Automacao_Email_Python\\{arquivo_nome}'
                planilha_anexo_modificado.to_excel(caminho_arquivo, index=False)
                st.session_state['caminho_arquivo'] = caminho_arquivo
                st.success(f"Anexo editado e salvo em {caminho_arquivo}")

    if 'caminho_arquivo' in st.session_state:
        st.subheader('Enviar Email')
        destinos = st.text_input("Digite o Email que será enviado:")
        if st.button('Enviar Email'):
            if destinos:
                email_criacao = Criacao_Email(usuario, senha, st.session_state['caminho_arquivo'], destinos, servico_email) #, st.session_state['servico_email'])
                msg = email_criacao.criacao_email()
                email_criacao.enviar_email(msg)
            else:
                st.error("Preencha o email de destino!")

if __name__ == "__main__":
    main()



