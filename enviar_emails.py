import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os

# Carregar a planilha
df = pd.read_excel('TESTE-EMAIL-AUTOMATICO-RV.xlsx', sheet_name='Página1')

# Configurações do e-mail
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'dados@fortsunbrasil.com'
smtp_password = 'xxii smwn jwhx vcsj'  # Use a senha de aplicativo aqui

# Função para criar o PDF com os dados do colaborador
def criar_pdf(nome, cust_id, status, cargo, admissao, cpf, supervisao):
    pdf_filename = f"{nome.replace(' ', '_')}_dados.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=A4)
    c.setFont("Helvetica", 12)
    c.drawString(100, 800, f"Dados do Consultor - {nome}")
    c.drawString(100, 780, f"Nome: {nome}")
    c.drawString(100, 760, f"Cust ID: {cust_id}")
    c.drawString(100, 740, f"Status: {status}")
    c.drawString(100, 720, f"Cargo: {cargo}")
    c.drawString(100, 700, f"Data de Admissão: {admissao}")
    c.drawString(100, 680, f"CPF: {cpf}")
    c.drawString(100, 660, f"Supervisão: {supervisao}")
    c.drawString(100, 640, "Atenciosamente,")
    c.drawString(100, 620, "Equipe de Dados da Fortsun")
    c.drawString(100, 600, "dados@fortsunbrasil.com")
    c.save()
    return pdf_filename

# Função para enviar o e-mail
def send_email(to_email, subject, body, pdf_filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Anexar o arquivo PDF
        with open(pdf_filename, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', f'attachment; filename="{pdf_filename}"')
            msg.attach(part)

        # Conectar ao servidor SMTP com TLS
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_email, msg.as_string())
            print(f'E-mail enviado com sucesso para {to_email}')

        # Remover o arquivo PDF após o envio
        os.remove(pdf_filename)

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

    # Criar o arquivo PDF
    pdf_filename = criar_pdf(nome, cust_id, status, cargo, admissao, cpf, supervisao)

    # Criar o corpo do e-mail
    subject = f'Dados do Consultor - {nome}'
    body = f"""
    Olá {nome},

    Seguem os dados do consultor em anexo.

    Atenciosamente,
    Equipe de Dados da Fortsun
    dados@fortsunbrasil.com
    ...
    """

    # Enviar o e-mail com o anexo PDF
    send_email(to_email, subject, body, pdf_filename)
