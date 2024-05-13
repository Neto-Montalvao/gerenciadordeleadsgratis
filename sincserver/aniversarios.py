import datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
import time
import os
import re
import bd

def encontrar_ultimo_arquivo(diretorio):
    arquivos_xlsx = [f for f in os.listdir(diretorio) if f.endswith('.xlsx') or f.endswith('.xls')]
    caminhos = [os.path.join(diretorio, basename) for basename in arquivos_xlsx]
    ultimo = max(caminhos, key=os.path.getmtime)
    return ultimo


hoje = datetime.now().date()
cor = {
    'Carol':{'email':'anacarolinagalvaomontalvao@gmail.com'},
    'Generozo':{'email':'geraldo@montalvaoimoveis.com.br'},
    'Joao':{'email':'joaomarcosleles01@gmail.com'},
    'Luiza':{'email':'luiza@montalvaoimoveis.com.br'},
    'Paula':{'email':'anapaula@montalvaoimoveis.com.br'},
    'Rosana':{'email':'rosana@montalvaoimoveis.com.br'},
    'Teste':{'email':'neto@montalvaoimoveis.com.br'}
}

# Informações da conta de e-mail de onde os e-mails serão enviados
email_envio = 'sac@montalvaoimoveis.com.br'
senha = ''

# Configurações do servidor de e-mail
servidor = smtplib.SMTP('smtp.gmail.com', 587)
servidor.starttls()
servidor.login(email_envio, senha)

def apenas_numeros(string):
    return re.sub(r'\D', '', str(string))
while True:
    hoje = datetime.now().date()
    df = pd.read_excel('gdrive\\Leads-Info\\!Geral\\alldados.xlsx')
    df2 = pd.read_excel(encontrar_ultimo_arquivo('tabelas\\'))


    df2['Telefone'] = df2['Telefone'].apply(apenas_numeros)
    df['Contato'] = df['Contato'].apply(apenas_numeros)

    for index, row in df.iterrows():

        if not pd.isnull(row['Data de nascimento do cliente']):

            data_nascimento = row['Data de nascimento do cliente'].date()
            
            if (data_nascimento.day == hoje.day) and (data_nascimento.month == hoje.month):
                print('aniversário de alguém')
                try:
                    bd.get(f'i{row['Contato']}2024')
                except:
                    bd.set(f'i{row['Contato']}2024', 'True')
                    msg = MIMEMultipart()
                    msg['From'] = email_envio
                    msg['To'] = cor[row['Corretor']]['email']
                    nome = df2.loc[df2['Telefone'] == row['Contato']]['Nome']
                    if not nome.empty:
                        nome = nome.iloc[0]
                    else:
                        nome = "desconhecido"

                    msg['Subject'] = f"Aniversário do cliente {nome}"
                    corpo_email = f"Hoje é o aniversário de {nome}. Entre em contato pelo número {row['Contato']}."
                    msg.attach(MIMEText(corpo_email, 'plain'))
                    texto = msg.as_string()
                    servidor.sendmail(email_envio, cor[row['Corretor']]['email'], texto)

                    corpo_email = f"for:{row['Corretor']} \n Hoje é o aniversário de {nome}. Entre em contato pelo número {row['Contato']}."
                    msg.attach(MIMEText(corpo_email, 'plain'))
                    texto = msg.as_string()

                    servidor.sendmail(email_envio, 'neto@montalvaoimoveis.com.br', texto)
                
    time.sleep(86400)