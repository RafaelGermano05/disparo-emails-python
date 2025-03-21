import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Carregar a planilha
df = pd.read_excel('TESTE-EMAIL-AUTOMATICO-RV.xlsx', sheet_name='Página1')

# Configurações do e-mail
smtp_server = 'smtp.gmail.com'
smtp_port = 587  # Porta para TLS
smtp_user = 'dados@fortsunbrasil.com'
smtp_password = 'xxii smwn jwhx vcsj'

# Função para enviar e-mail
def send_email(to_email, subject, body):
    try:
        # Configuração da mensagem
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Conectar ao servidor SMTP com TLS
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_email, msg.as_string())
            print(f'E-mail enviado com sucesso para {to_email}')

    except smtplib.SMTPAuthenticationError:
        print(f'Erro de autenticação ao enviar e-mail para {to_email}. Verifique as credenciais.')
    except smtplib.SMTPConnectError:
        print(f'Erro de conexão com o servidor SMTP.')
    except smtplib.SMTPRecipientsRefused:
        print(f'O destinatário {to_email} foi recusado.')
    except smtplib.SMTPException as e:
        print(f'Erro ao enviar e-mail para {to_email}: {e}')

# Iterar sobre as linhas da planilha
for index, row in df.iterrows():
    to_email = row['E-mail FAST PE \n Parceira Estratégica']
    cust_id = row['Cust Id FAST']
    nome = row['Nome']
    status = row['Status']
    cargo = row['Cargo']
    admissao = row['Admissão']
    cpf = row['CPF']
    supervisao = row['Supervisão']

    # Criar o corpo do e-mail
    subject = f'Dados do Consultor - {nome}'
    body = f"""
    Olá,

    Seguem os dados do consultor:

    Nome: {nome}
    Cust ID: {cust_id}
    Status: {status}
    Cargo: {cargo}
    Data de Admissão: {admissao}
    CPF: {cpf}
    Supervisão: {supervisao}

    Atenciosamente,
    Equipe de RH
    """

    # Enviar o e-mail
    send_email(to_email, subject, body)
