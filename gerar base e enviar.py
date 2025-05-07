import oracledb
import pandas as pd
import subprocess #para atualizar o projeto
import win32com.client # para atualizar planilha
import win32com.client as win32 
import xlwings as xw
import os # para navegar no windows
# Para notificar por email
import time
from datetime import datetime
import shutil
from PIL import ImageGrab
import tempfile
import smtplib
from email.message import EmailMessage
import csv

email_para = ['informar Email']
email_cc = ['informar Email']
assunto = 'Informar o assunto'

# Inicializa o modo thick com o client Oracle
oracledb.init_oracle_client(lib_dir=r"instantclient_23_8")  # ajuste o caminho conforme sua máquina


username = 'xxxxxx'
password = 'xxxxxxxx'
host = 'xxxxxxxxx'
port = xxxxxxx
service_name = 'xxxxxx'

# Construir DSN corretamente
dsn = oracledb.makedsn(host, port, service_name=service_name)

query = """
-- SUA CONSULTA AQUI
    select
        *
    from nome_da_tabela
"""

arquivo = r'endereço da pasta'

try:
    connection = oracledb.connect(user=username, password=password, dsn=dsn)
    print("Conexão bem-sucedida!")

    df = pd.read_sql(query, con=connection)
    df.to_excel(arquivo, index=False, engine='openpyxl')  # ✅ usando o caminho correto
    print("Consulta executada com sucesso!")
    print(df)

except Exception as e:
    print(f"Erro ao gerar a planilha: {e}")
    
    
    
Assinatura = 'xxxxxx'
Cargo = 'xxxxxxxx'
Setor = 'xxxxxxxxxxxxxx'
Gerencia = 'xxxxxxxxxxxxxxxx'
contato = 'xxxxxxxxxxxxxxx'
Link = 'xxxxxxxxxxxxxx'
email = 'xxxxxxxxxxxxxxxxxxxxx'


base = r'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'   


# Criando e-mail no Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# Defina o destinatário
# mail.To = "; ".join(Email_Teste)  # Para testes
mail.To = "; ".join(email_para)  # Certifique-se de que está correto
mail.cc = "; ".join(email_cc)

# Obter data e hora formatadas
data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")

mail.Subject = f"{assunto} - {data_hora}"  # Lembre de definir o assunto

corpo = f"""
<html>
    <body>
        <p>Prezados(a), espero que estejam bem.</p>
        <p>
            Encaminho a base abaixo. Solicito, por gentileza, que analisem.             
        </p>
        <img src="cid:screenshot" alt="Base_conexÃµes" style="width:100%; max-width:600px;">
        <p>Agradeço pela atenção.</p>
        <p><b>Atenciosamente,</b></p>
        <table>
            <tr>
                <td>
                    <img src="{r'C:\\img\logo Assinatura.png'}"
                         alt="Logo Equatorial Energia" style="height:50px; width:100px;">
                </td>
                <td style="padding-left:10px; font-size:12px; font-family:Arial, sans-serif;">
                    <b>{Assinatura}</b><br>
                    {Cargo}<br>
                    {Setor}<br>
                    {Gerencia}<br>
                    <b>Whatsapp:</b> <a href="{Link}">{contato}</a><br>
                    <b>Email:</b> <a href="mailto:{email}">{email}</a>
                </td>
            </tr>
        </table>
    </body>
</html>
"""
mail.HTMLBody = corpo
attachment = base


# Verifica se o arquivo existe antes de anexar
if os.path.exists(attachment):
    mail.Attachments.Add(attachment)
else:
    print(f"⚠️ Erro: O arquivo {attachment} não foi encontrado!")

# Verifica se o destinatário foi definido antes de enviar
if not mail.To:
    print("⚠️ Erro: Nenhum destinatário definido!")
else:
    mail.Send()
    print("✅ E-mail enviado com sucesso!")
    
    
# Caminho da pasta
pasta_prints = r'endereço da pata da base'

# Verifica se a pasta existe
if os.path.exists(pasta_prints):
    # Remove todos os arquivos dentro da pasta
    for arquivo in os.listdir(pasta_prints):
        caminho_arquivo = os.path.join(pasta_prints, arquivo)
        try:
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)  # Apagar arquivos
            elif os.path.isdir(caminho_arquivo):
                shutil.rmtree(caminho_arquivo)  # Apagar subpastas (se houver)
        except Exception as e:
            print(f"Erro ao excluir {arquivo}: {e}")

    print("✅ Todos os arquivos foram apagados da pasta 'prints'.")
else:
    print("⚠️ A pasta 'prints' não existe.")


