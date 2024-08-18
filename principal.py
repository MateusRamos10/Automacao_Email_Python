# Ler Email
from imap_tools import MailBox, AND
from pandas import DataFrame as df
import pandas as pd
from io import BytesIO
from datetime import datetime
# Enviar Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Classe Encontrar_Email
class Encontrar_Email:
    def __init__(self, usuario, senha):
        self.usuario = usuario
        self.senha = senha
        self.meu_email = MailBox("imap.gmail.com").login(usuario, senha)

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

        return resultado

    def edita_anexo(self, anexo_escolhido):
        dados_em_bytes = anexo_escolhido
        buffer = BytesIO(dados_em_bytes)
        planilha_anexo = pd.read_excel(buffer)
        planilha_anexo_modificado = planilha_anexo.drop(['ASDAD', 'ASDASD', 'ASDASD.1', 'ASDSAD'], axis=1)
        return planilha_anexo_modificado

# Classe Criacao_Email
class Criacao_Email:
    def __init__(self, usuario, senha, caminho_arquivo, destinos):
        self.usuario = usuario
        self.senha = senha
        self.caminho_arquivo = caminho_arquivo
        self.destinos = destinos

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
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.starttls()
            s.login(self.usuario, self.senha)
            s.send_message(msg)
        
        print('Email Enviado')

# Função main
def main():
    usuario = input("Digite seu email: ")
    senha = input("Digite sua senha: ")

    email_obj = Encontrar_Email(usuario, senha)

    destinatario = input("Digite o Email do Destinatário: ")
    assunto = input("Digite o assunto para filtrar: ")

    lista_cabecalho = email_obj.lista_emails(destinatario, assunto)
    resultado = email_obj.cria_dataframe(lista_cabecalho)
    print(resultado)

    email_escolhido = int(input("Digite o código do email escolhido: "))
    if 0 <= email_escolhido < len(lista_cabecalho):
        anexo_escolhido = lista_cabecalho[email_escolhido][3]

    planilha_anexo_modificado = email_obj.edita_anexo(anexo_escolhido)
    data_atual = datetime.now().strftime('%d-%m-%Y')
    arquivo_nome = f'arquivo{data_atual}.xlsx'
    caminho_arquivo = f'C:\\GitHub\\Automacao_Email_Python\\{arquivo_nome}'
    planilha_anexo_modificado.to_excel(caminho_arquivo, index=False)

    destinos = input("Digite o Email que será enviado: ")
    email_criacao = Criacao_Email(usuario, senha, caminho_arquivo, destinos)

    msg = email_criacao.criacao_email()
    email_criacao.enviar_email(msg)

if __name__ == "__main__":
    main()
