# %%
import os
from datetime import datetime
import time
import pandas as pd
from copy import copy
from pathlib import Path
import re
import xlwings as xw
import numpy as np
import sys
import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account


# %%
#connect to google drive, get and write data


# %%
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY ={
    "type": "service_account",
    "project_id": st.secrets["project_id"],
    "private_key_id": st.secrets["private_key_id"],
    "private_key": st.secrets["private_key"].replace('\\n', '\n'),
    "client_email": st.secrets["client_email"],
    "client_id": st.secrets["client_id"],
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://accounts.google.com/o/oauth2/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": st.secrets["client_x509_cert_url"],
    "universe_domain": "googleapis.com"
}
# Escribe aquí el ID de tu documento:
SPREADSHEET_ID = '1LdjjQh4IX_lGM9iup3FXOeWxqESBIeVwDa27uPoKcVI'
RANGE_NAME_GET="Fill F!A:O"

creds = None
creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)

service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# Llamada a la api
result = (
        sheet.values()
        .get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME_GET)
        .execute()
    )
# Extraemos values del resultado
values = result.get('values',[])
df=pd.DataFrame(values[1:],columns=values[0])
df_fix=df[df["Fecha"] != '']
df_index=df_fix.index
maxrow=df_index[-1]+3


# %%
#display(df_fix)


# %%
categorias=(
        sheet.values()
        .get(spreadsheetId=SPREADSHEET_ID, range='Config!L17:V')
        .execute()
    ).get('values',[])
categorias=pd.DataFrame(categorias[1:],columns=categorias[0])


# %%
RANGE_NAME_FILL=f"Fill F!B{maxrow}"


# %%
def llenar(valores):
    filling = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME_FILL,
        valueInputOption='USER_ENTERED',
        body={'values': [valores]}
    )
    return filling.execute()


# %%
#streamlit code


# %%


# %%
fecha = st.date_input('Selecciona una fecha', datetime.now(),format='MM/DD/YYYY',key='fecha')

tipo=st.radio('Tipo',['Ingresos','Egresos'],key='tipo')

CF=st.toggle('Cashflow')

if CF==True:
    CF_l='Cashflow'
else:
    CF_l='Nocashflow'


TDC=st.toggle('Tarjeta de Credito')

if TDC==True:
    TDC_l='Si'
else:
    TDC_l='No'


monto=st.number_input('Monto',min_value=0.00,format='%.2f',key='monto')

categoria=st.selectbox('Categoria',categorias.columns,key='cat')

subcategoria=st.selectbox('Subcategoria',categorias.loc[~categorias[categoria].isin(['', None]),categoria],key='subcat')

selected_option_distrito = st.selectbox('Distrito', df_fix.loc[~df_fix['Distrito'].isin(['', None]),'Distrito'].drop_duplicates( keep='first').tolist() + ['CUSTOM'],key='distr')

if selected_option_distrito == 'CUSTOM':
    custom = st.text_input('Ingreso el nuevo distrito:')
    distrito=custom
else:
    distrito=selected_option_distrito

selected_option_establecimiento = st.selectbox('Establecimiento', df_fix.loc[~df_fix['Establecimiento'].isin(['', None]),'Establecimiento'].drop_duplicates( keep='first').tolist() + ['CUSTOM'],key='establ')

if selected_option_establecimiento == 'CUSTOM':
    custom = st.text_input('Ingreso el nuevo establecimiento:')
    establecimiento=custom
else:
    establecimiento=selected_option_establecimiento

selected_option_item = st.selectbox('Item', df_fix.loc[~df_fix['Item'].isin(['', None]),'Item'].drop_duplicates( keep='first').tolist() + ['CUSTOM'],key='item')

if selected_option_item == 'CUSTOM':
    custom = st.text_input('Ingreso el nuevo distrito:')
    item=custom
else:
    item=selected_option_item


st.write(f'{fecha} {tipo} {CF_l} {TDC_l} {monto} {categoria} {subcategoria} {distrito} {establecimiento} {item}')

valores=[fecha.strftime('%d-%b-%Y'),tipo,distrito,establecimiento,categoria,subcategoria,item,monto,CF_l,TDC_l]

boton=st.button('FILL', on_click=llenar, args=(valores,))

if boton==True:
    filling2 = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="DB F!A2",
        valueInputOption='USER_ENTERED',
        body={'values': (sheet.values().get(spreadsheetId=SPREADSHEET_ID,range=f'Fill F!A2:O{maxrow-1}').execute()).get('values',[])}
    ).execute()


# %%
on = st.toggle('Mostrar ultimos 10 gastos')

if on:
    st.table(df_fix[-10:])





