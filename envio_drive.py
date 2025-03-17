import os
import datetime
from dotenv import load_dotenv
from googleapiclient.http import MediaFileUpload
from servico_google import acessando_drive


load_dotenv()
PASTA_RAIZ_ID = os.getenv('PASTA_DRIVE_LEVANTAMENTO_DEBITO_GUARULHOS')

def compartilhar_pasta(drive_service, pasta_id, seu_email):
    permissão = {
        'type': 'user',  
        'role': 'writer',  
        'emailAddress': seu_email
    }
    drive_service.permissions().create(
        fileId=pasta_id,
        body=permissão,
        fields='id'
    ).execute()

def salvar_drive(nome_arquivo, resp):
    mes_atual = datetime.datetime.now().strftime("%m/%Y")
    nome_pasta = f'Levantamento de debito: {mes_atual}'
    
    drive = acessando_drive()
    response = drive.files().list(
        q=f"name = '{nome_pasta}' and '{PASTA_RAIZ_ID}' in parents and mimeType = 'application/vnd.google-apps.folder'",
        spaces='drive'
    ).execute()
    pastas = response.get('files', [])
    
    if pastas:
        pasta_id = pastas[0]['id']
    else:
        folder_metadata = {
            'name': nome_pasta,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [PASTA_RAIZ_ID]
        }
        pasta = drive.files().create(body=folder_metadata, fields='id').execute()
        pasta_id = pasta['id']
        seu_email = "inov2@controller-oraculus.com.br"  
        compartilhar_pasta(drive, pasta_id, seu_email)  
        print(f"Pasta criada/encontrada: https://drive.google.com/drive/folders/{pasta_id}")
    
    response = drive.files().list(
        q=f"name = '{resp}' and '{pasta_id}' in parents and mimeType = 'application/vnd.google-apps.folder'",
        spaces='drive'
    ).execute()
    subpastas = response.get('files', [])
    
    if subpastas:
        subpasta_id = subpastas[0]['id']
    else:
        subfolder_metadata = {
            'name': resp,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [pasta_id]
        }
        subpasta = drive.files().create(body=subfolder_metadata, fields='id').execute()
        subpasta_id = subpasta['id']
    
    file_metadata = {
        'name': nome_arquivo,
        'parents': [subpasta_id]
    }
    media = MediaFileUpload(nome_arquivo, resumable=True)
    arquivo = drive.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"Arquivo salvo em: https://drive.google.com/file/d/{arquivo['id']}")
    
    return subpasta_id
    

        