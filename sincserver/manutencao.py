import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import locale
import requests
import json
import bd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from datetime import timedelta
import re
from urllib.parse import quote


# Suponha que você tenha um DataFrame chamado df
df = pd.read_excel('H:\\Outros computadores\\ImobServer\\sincserver\\tabelas\\villagio_real_3_26_03_2024 (100).xlsx')

dfd = df.copy()
dfd['Data de Cadastro'] = pd.to_datetime(dfd['Data de Cadastro'], format='%d/%m/%Y - %H:%M:%S')

dfd.sort_values('Data de Cadastro', ascending=False, inplace=True)

df_duplicates = dfd[dfd.duplicated(subset=['E-mail'], keep=False) | dfd.duplicated(subset=['Telefone'], keep=False)]

# Cria uma máscara booleana que identifica as linhas duplicadas dentro de um intervalo de 48 horas
mask = dfd.groupby(['E-mail', 'Telefone'])['Data de Cadastro'].transform(lambda x: x.diff().dt.total_seconds() > -48*3600)
# Remove as linhas duplicadas
df_duplicates = dfd[mask]
print(df_duplicates)
df_duplicates.to_excel('c:/wppautomsgs/aiaiaia.xlsx')
