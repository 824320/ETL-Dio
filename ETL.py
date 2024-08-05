
import pandas as pd
import openai


# Carregar os dados da planilha Excel
df = pd.read_excel(r'caminho do arquivo excel')

# Supondo que as colunas sejam 'Nome', 'Email' e 'Celular'
clientes = df[['Nome', 'Celular', 'Email']]

# Configurar a chave da API da OpenAI
openai.api_key = ' sua chave openai'

def criar_mensagem_gpt(nome):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "Você é um assistente que cria mensagens de bom dia."},
            {"role": "user", "content": f"Crie uma mensagem de bom dia personalizada para {nome}."}
        ]
    )
    mensagem = response.choices[0].message['content'].strip()
    return mensagem

# Adicionar a coluna de Mensagens ao DataFrame usando o GPT
clientes['Mensagem'] = clientes['Nome'].apply(criar_mensagem_gpt)

# Verificar se as mensagens foram geradas corretamente
#print(clientes.head())

# segunda parte do codigo enviar e-mail


import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def enviar_email(destinatario, mensagem):
    remetente = "seu email"
    senha = "sua chave de email app google"

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = "Bom Dia!"

    msg.attach(MIMEText(mensagem, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        text = msg.as_string()
        server.sendmail(remetente, destinatario, text)
        server.quit()
        print(f'E-mail enviado para {destinatario}')
    except Exception as e:
        print(f'Falha ao enviar e-mail para {destinatario}: {e}')

# Enviar e-mails para todos os clientes
clientes.apply(lambda x: enviar_email(x['Email'], x['Mensagem']), axis=1)


