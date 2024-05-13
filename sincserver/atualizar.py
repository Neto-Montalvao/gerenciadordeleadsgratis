from openpyxl import load_workbook
import re
from datetime import timedelta
import datetime
import locale
import json
import requests
import bd

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

import os
import time
import pandas as pd
import autoenv as at
import subprocess
import webbrowser as web
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

subprocess.Popen(['python', "C:\\wppautomsgs\\main.py"])
subprocess.Popen(['python', "C:\\sincserver\\verifica.py"])
subprocess.Popen(['python', "C:\\sincserver\\verifica2.py"])
subprocess.Popen(['python', "C:\\sincserver\\aniversarios.py"])


def verificar_arquivo():
    #try:
        forca = True
        path = encontrar_ultimo_arquivo('c:\\sincserver\\tabelas\\', 'villagio_real_3_')

        antigo = pd.read_excel("c:\\sincserver\\recontatos.xlsx")
        print('vai')
        web.open(f"https://hub.exent.com.br/projects/99/leads", new=1, autoraise=True)
        while True:
            time.sleep(65)
            novo = pd.read_excel(encontrar_ultimo_arquivo('c:\\sincserver\\driveautom\\', 'Recontato - Villagio Real 3 (respostas)'))

            if not novo.equals(antigo) or forca:
                enviar = False
                antigo['Carimbo de data/hora'] = pd.to_datetime(antigo['Carimbo de data/hora'], format='%d/%m/%Y')
                novo['Carimbo de data/hora'] = pd.to_datetime(novo['Carimbo de data/hora'], format='%d/%m/%Y')

                antigo = antigo.sort_values(by=['Carimbo de data/hora']).reset_index(drop=True)
                novo = novo.sort_values(by=['Carimbo de data/hora']).reset_index(drop=True)

                diferentes = pd.concat([antigo, novo]).drop_duplicates(keep=False)

                df = pd.read_excel('allleads.xlsx')

                def apenas_numeros(string):
                    return re.sub(r'\D', '', f'{string}')
                
                diferentes['Número'] = diferentes['Número'].apply(apenas_numeros)
                
                df['tell'] = df['tell'].apply(apenas_numeros)

                if not diferentes.empty or forca:
                    for i, r in diferentes.iterrows():
                        df.loc[(df['corretor'] == 'em analise') & ((df['tell'] == r['Número']) | (df['id'] == r['ID'])), 'corretor'] = r['corretor repassado']
                        
                    envios = [
                        {'nome': 'neto', 'tell': '5514997603977', 'msg': 'enviandocobrancaatrasada', 'imgc': 'vazia', 'fecha':True},
                        {'nome': 'neto', 'tell': '5514997603977', 'msg': 'enviandocobrancaatrasada', 'imgc': 'vazia', 'fecha':True}
                    ]
                    
                    dfa = df.copy()

                    dfa['data de repasse'] = pd.to_datetime(dfa['data de repasse'], format='%d/%m/%Y - %H:%M:%S')
                    rec = eval(bd.get('recontatos'))
                    dfa = dfa[(dfa['corretor'] == 'em analise') & (~dfa['tell'].isin(rec)) & (dfa['data de repasse'] < (datetime.datetime.now() - timedelta(days=2)))]

                    

                    for i, r in dfa.iterrows():
                        enviar = True
                        envios.append({'nome': 'Montalvao Imoveis', 'tell': '5514996892957', 'msg': f'''Passaram-se mais de 2 dias desde o contato com {r["nome"]} - ID: {r['id']} - Tell: {re.sub(r'\D', '', r["tell"])}.\n por favor, reporte o status no formulário''', 'imgc': 'vazia', 'fecha':True})
                        rec.append(r['tell'])

                    bd.set('recontatos', f'{rec}')

                    novo.to_excel("c:\\sincserver\\recontatos.xlsx", index=False)

                    df.to_excel('c:\\sincserver\\allleads.xlsx', index=False)


                    book = load_workbook('c:\\sincserver\\allleads.xlsx')
                    sheet = book.active
                    for column_cells in sheet.columns:
                        length = max(len(str(cell.value)) for cell in column_cells)
                        sheet.column_dimensions[column_cells[0].column_letter].width = length+5

                    book.save('c:\\sincserver\\allleads.xlsx')


                    if enviar:
                        enviar = False
                        url = 'http://localhost:5000/data'
                        headers = {'Content-Type': 'application/json'}

                        data = json.dumps(envios)

                        response = requests.post(url, headers=headers, data=data)

                        if response.status_code == 200:
                            print(response.json())
                        else:
                            print(f'Error: {response.status_code}')

            path = encontrar_ultimo_arquivo('c:\\sincserver\\tabelas\\', 'villagio_real_3_')

            dados_atuais = pd.read_excel(path)

            dados_atuais = dados_atuais[dados_atuais['ID'] > int(bd.get('last'))]

            if not dados_atuais.empty:
                forca = False
                at.enviasincserver(pd.read_excel(path))
                print('sim')
            else:
                print('nao')
    #except Exception as e:
   #     send_email(f"Erro no sistema", f'{e}', "neto@montalvaoimoveis.com.br")

def encontrar_ultimo_arquivo(diretorio, parte):
    arquivos_xlsx = [f for f in os.listdir(diretorio) if (f.endswith('.xlsx') or f.endswith('.xls')) and parte in f]
    caminhos = [os.path.join(diretorio, basename) for basename in arquivos_xlsx]
    ultimo = max(caminhos, key=os.path.getmtime)
    deletar_arquivos(diretorio, parte)
    return ultimo

def deletar_arquivos(diretorio, parte):
    arquivos_xlsx = [f for f in os.listdir(diretorio) if (f.endswith('.xlsx') or f.endswith('.xls')) and parte in f]
    
    caminhos = [(os.path.join(diretorio, basename), os.path.getmtime(os.path.join(diretorio, basename))) for basename in arquivos_xlsx]
    
    caminhos.sort(key=lambda x: x[1])
    
    if len(caminhos) > 2:
        for i in range(len(caminhos) - 2):
            try:
                os.remove(caminhos[i][0])
            except:
                None




def send_email(subject, body, to):
    from_address = "sac@montalvaoimoveis.com.br"
    password = ""

    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_address, password)
    text = msg.as_string()
    server.sendmail(from_address, to, text)
    server.quit()


verificar_arquivo()