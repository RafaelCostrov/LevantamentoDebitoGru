import os
import base64
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


load_dotenv()
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')
SCOPES = os.getenv('SCOPES_EMAIL').split(',')  
EMAIL_USER = os.getenv('EMAIL_USER')
TEMPLATE_PATH = os.getenv('TEMPLATE_PATH')

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
).with_subject(EMAIL_USER)
service = build('gmail', 'v1', credentials=credentials)

def carregar_template(nome, link):
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as file:
        template = file.read()
        return template.replace("{0}", nome).replace("{1}", link)

def criar_email(destinatario, assunto, nome, link):
    email = MIMEMultipart()
    email['to'] = destinatario
    email['from'] = EMAIL_USER
    email['subject'] = assunto
    email.attach(MIMEText(carregar_template(nome, link), 'html'))
    
    raw_message = base64.urlsafe_b64encode(email.as_bytes()).decode('utf-8')
    return {'raw': raw_message}

def enviar(destinatario, assunto, nome, link):
    email = criar_email(destinatario, assunto, nome, link)
    service.users().messages().send(
        userId='me',  
        body=email
    ).execute()
    print(f"E-mail enviado com sucesso para {destinatario}!")
