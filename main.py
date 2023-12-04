# Programa para calcular el due-date de las ordenes
# Modificado Agosto 23/2023
import tkinter as tk
from tkinter import messagebox
from tkinter import *
import csv
import mimetypes
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import datetime  # as dt
from datetime import date
from datetime import datetime as dt
import time
from tkinter.constants import *
from tkcalendar import Calendar
import pytz
import locale
from oauth2client.service_account import ServiceAccountCredentials
#from Fun_PromesaCiente import *
from datetime import datetime, timedelta
import pandas as pd                                                                            
import DespliegueDB
from google.auth.transport.requests import Request
import gspread
import warnings
import threading
import gdown
import io
import os
from urllib.request import urlopen
from urllib.parse import urlparse
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
import smtplib, ssl
from googleapiclient.http import MediaFileUpload
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import base64
import re
import requests
from bs4 import BeautifulSoup
import oracledb

sess = requests.Session()
adapter = requests.adapters.HTTPAdapter(max_retries=20)
sess.mount('http://', adapter)

user = "ADMIN"
dsn = "tapreturnsdb_high"
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

file_id0 = "1F0L_aHVNNhGuV-KNnuT6nCr_X1Af3l3E"  # cortes2023.xlxs
file_id1 = "15vHlzGFgi9MjxyclqmNArvheijJhLSK5"  # tiempos.xls

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send'
]
# style = ttk.Style()
# print(style.theme_use())
# selected = ""
pw = "pROCESOS2023"
wallet_pw = "!Tr3svecestres!"
con = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                       wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)
print("Database version", con.version)
# con.close()
proxies = {'https': 'http:tapaserver.dyndns.org'}
credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json",
                                                               scopes)  # access the json key you downloaded earlier
file = gspread.authorize(credentials)  # authenticate the JSON key with gspread
ss = file.open("EficienciaReporte")
service = build('drive', 'v3', credentials=credentials)
creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', scopes)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', scopes)
        creds = flow.run_local_server(port=0)

    with open('token.json', 'w') as token:
        token.write(creds.to_json())

services = build('gmail', 'v1', credentials=creds)
results = services.users().labels().list(userId='me').execute()
url3 = "https://www.banxico.org.mx/SieInternet/consultarDirectorioInternetAction.do?sector=6&accion=consultarCuadro&idCuadro=CF102&locale=es"
html = sess.get(url3).content
df_list = pd.read_html(html)
df = df_list[0]

resultWeekend = df.iloc[8, 2]
port = 587  # For starttls
smtp_server = "smtp.gmail.com"
sender_email = "tapatio.procesos@gmail.com"
receiver_email = "juan.a.o.delgado@gmail.com"
receiver = "juan.a.o.delgado@gmail.com"

detener_thread = False
detener_thread2 = False
i = 1
super_user = 0
result = ""
result1 = ""
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
DOF = url = URL = page = soup = texto = ""
cambio = 0
entry1 = 0
Tipo_Cambio = 0.0
Garantia90 = ['300 ENGINE ASSEMBLY', '400 Transmiss,Transaxle', '412 Transfer Case Assy', '600 BATTERY']
Nogarantia = ['590 Eng/Motor Cont Mod', '591 Chassis Cont Mod', '337 Throttle Body Assy']


def abrir_conexion():
    global con

    try:
        cursor = con.cursor()
        cursor.execute("SELECT 1 FROM dual")
    except oracledb.Error as e:
        con = oracledb.connect(user=user, password=pw, dsn=dsn, config_dir=r"C:\\network\\admin",
                               wallet_location=r"C:\\network\\admin", wallet_password=wallet_pw)
        print("Database version1", con.version)


def scrap_web():
    global DOF, Tipo_Cambio, result, result1, diff, i

    url = service.files().get_media(fileId="18F4Ix9C_2q7tjimDycg6v7XnCAb258FM").execute()
    URL = url.decode('utf-8')
    unverified_context = ssl._create_unverified_context()
    page = urlopen(URL, context=unverified_context)
    url = service.files().get_media(fileId="1Y-5sSIKrF1HmV58gPJpTWZE7yB4Qa9Ee").execute()
    diff = float(url.decode('utf-8'))
    page = page.read().decode("utf-8")
    soup = BeautifulSoup(page, "html.parser")
    texto = soup.get_text()
    if datetime.today().weekday() != 6 and datetime.today().weekday() != 5:
        resultado = re.search('DOLAR (.*\d)UDIS', texto)
        if resultado is None:
            result = resultWeekend
            date = datetime.today() - timedelta(days=3)
            result1 = date.strftime("%Y/%m/%d")
        else:
            result = resultado.group(1)
            resultado1 = re.search('Tipo de Cambio y Tasas al (.*)', texto)
            result1 = resultado1.group(1)
    elif datetime.today().weekday() == 6 or datetime.today().weekday() == 5:
        result = resultWeekend
        resultado1 = re.search('Tipo de Cambio y Tasas al (.*)', texto)
        result1 = resultado1.group(1)
        date = dt.strptime(result1, "%d/%m/%Y")
        if datetime.today().weekday() == 6:
            date = date - timedelta(days=2)
        elif datetime.today().weekday() == 5:
            date = date - timedelta(days=1)
        result1 = date.strftime("%Y/%m/%d")
    Tipo_Cambio = float(result)
    DOF = "Tipo de Cambio y Tasas al " + result1
    i = i + 1
    print(datetime.today().strftime("%H:%M:%S"))

scrap_web()
file_url = service.files().get(fileId=file_id0, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)
filename = f"cortes2023.xlsx"
try:
    request = service.files().get_media(fileId=file_id0)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")
file_url = service.files().get(fileId=file_id1, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)
filename = f"tiempos.xls"
try:
    request = service.files().get_media(fileId=file_id1)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")

url0 = "https://drive.google.com/file/d/11FW_HPRLaR-h2bk9eMs5VDTQvJ8Sg3B4/view?usp=share_link"  # festivos 2023
url = "https://drive.google.com/file/d/1CgwK7OCtiTXggOZpw3E39PgMlPEQo8WP/view?usp=share_link"  # claves.csv
url1 = "https://drive.google.com/file/d/1zTGwpZEryABsYqAf0xGISu1rYZ9MNTsF/view?usp=sharing"  # partes exportacion
url2 = "https://drive.google.com/file/d/1hwL7ujQ4rCTtOzwIQ1PQLF52GKM_Ywp3/view?usp=sharing"  # codigos partes
url3 = "https://drive.google.com/file/d/1gsiWvqon0f5zSGL8Xa9o27E46OU59O7_/view?usp=sharing"  # Folder ID
url14 = "https://docs.google.com/spreadsheets/d/10oVEZ0K7ZQp7FJL3e0QH_HjLAyr4NADcPfDXBBb9F6E/edit?usp=sharing"  # Listado regresos

output_path = 'claves.csv'
gdown.download(url, output_path, quiet=False, fuzzy=True)
output_path = 'festivos2023.csv'
gdown.download(url0, output_path, quiet=False, fuzzy=True)
output_path = 'tipopartes.csv'
gdown.download(url2, output_path, quiet=False, fuzzy=True)
output_path = 'foldersID.csv'
gdown.download(url3, output_path, quiet=False, fuzzy=True)

with open(r'claves.csv', mode='r') as infile:
    reader = csv.reader(infile)
    with open('claves_new.csv', mode='w') as outfile:
        writer = csv.writer(outfile)
        users = {rows[0]: rows[1] for rows in reader}

USA = 0
MEXICO = 12
MEXICO2 = 2
USA2 = 3
dfFest = pd.read_csv(r'festivos2023.csv')
ds = pd.read_csv(r'tipopartes.csv')

tipo_partes = ds['tipo de parte'].tolist()
tipo_partes.sort()

df = pd.read_excel(r'Tiempos.xls')
dc = pd.read_excel(r'Cortes2023.xlsx')
folder_id = ""
folder = ""
df1 = pd.DataFrame({
    "User": '',
    "Order #": '',
    "Part Store #": '',
    "Route": '',
    "Created": '',
    "Due Date part": '',
    "Due Date Order": '',
    "DueDate change": '',
    "Reason": '',
}, index=["Dummy"])
# print(dfFest)
dfpartes = []
regresosOK = []
regresosNG = []
media_paths = []
input_entries = []
QtyRtnNG = 0
QtyRtnPending = 0
QtyRtnaut = 0
festivosusa = dfFest["fechaUSA"].tolist()
festivosmex = dfFest["fechaMEX"].tolist()
festivosusa1 = dfFest["fechaUSAd"].tolist()
festivosmex1 = dfFest["fechaMEXd"].tolist()
output_path = 'PartesExportacion.csv'
df = df.set_index('Store')
dc = dict(dc.set_index('DIA').groupby(level=0). \
          apply(lambda x: x.to_dict(orient='list')))


def export_file():
    global credentials, dfpartes
    # Abrir la hoja de cálculo
    file = gspread.authorize(credentials)
    ss1 = file.open("Listado de partes Importacion-Exportacion")
    worksheet = ss1.get_worksheet(0)
    values = worksheet.get_all_values()
    dfpartes = pd.DataFrame(values[1:], columns=values[0])
    dfpartes = dfpartes.drop(dfpartes.columns[0], axis=1)
    dfpartes['No Parte'] = dfpartes['No Parte'].astype(str)
    dfpartes.to_excel("archivo.xlsx", index=False)
    print('Archivo descargado y guardado como archivo.xlsx')

current_time = datetime.now()
export_file()
control = 1


def listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button):
    global QtyRtnNG, QtyRtnPending, QtyRtnaut, con
    status_list = ['PendingDT', 'PendingAUT', 'DoneNG']
    abrir_conexion()

    for status in status_list:
        try:
            cursor = con.cursor()
            if selected in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO", "EMMANUEL LOPEZ"]:
                cursor.execute("SELECT COUNT(*) FROM regresos WHERE ESTATUS_REGRESO = :1", (status,))
            else:
                cursor.execute("SELECT COUNT(*) FROM regresos WHERE SELLER = :1 AND ESTATUS_REGRESO = :2",
                               (selected, status))
            if status == 'PendingDT':
                QtyRtnaut = cursor.fetchone()[0]
            elif status == 'PendingAUT':
                QtyRtnPending = cursor.fetchone()[0]
            elif status == 'DoneNG':
                QtyRtnNG = cursor.fetchone()[0]

        except oracledb.Error as error:
            print(f"Error de Oracle: {error}")
        finally:
            a = 0
    if QtyRtnPending:
        RtPnd_button.config(text=QtyRtnPending)
        RtPnd_button.config(state=tk.NORMAL)
    elif QtyRtnPending == 0:
        RtPnd_button.config(state=tk.DISABLED)
    if QtyRtnaut:
        RtPndaut_button.config(text=QtyRtnaut)
    if QtyRtnNG:
        Rtng_button.config(text=QtyRtnNG)

def mantener_conexion_activa():
    cursor = con.cursor()
    cursor.execute("SELECT 1 FROM dual")
    cursor.close()
def main_program(super_user, hoja, listado):
    with open('foldersID.csv', mode='r') as file:
        IDfolder = csv.reader(file)
        next(IDfolder)
        for row in IDfolder:
            if row[0] == selected:
                folder_id = row[1]
                break

    def actualizacion():
        global DOF, resultWeekend, diff, result1, result, control, Tipo_Cambio, detener_thread

        while not detener_thread:
            export_file()
            scrap_web()
            if control == 1:
                T_Cambio = tk.Label(main_window, font=("times", 13, "bold"), fg="green",
                                    text=f"$ {Tipo_Cambio}")
            else:
                T_Cambio = tk.Label(main_window, font=("times", 13, "bold"), fg="red",
                                    text=f"$ {Tipo_Cambio}")
            T_Cambio.place(x=160, y=37)
            control = 1 - control

            for segundos in range(2 * 60 * 60):
                if detener_thread:
                    print("Hilo terminado thread")
                    break
                time.sleep(1)
    def conexion():
        global detener_thread2
        while not detener_thread2:
            listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button)

            for segundos in range(30):
                if detener_thread2:
                    print("Hilo terminado thread2")
                    break
                time.sleep(1)
    def add_action():

        selected_radio = var.get()
        selected_combobox = combobox.get()
        if selected_radio in [1, 2, 3, 6, 7, 8] and selected_combobox == "EBAY TJ":
            messagebox.showwarning("Aerta", "No se puede usar ruta EBAY TJ si la parte esta en USA ")
        elif selected_radio == 2 and selected_combobox == "WILL CALL 1":
            messagebox.showwarning("Alerta", "No se envia a WILL CALL1 desde Nirvana ")
        elif selected_radio == 1 and selected_combobox == "WILL CALL 1":
            messagebox.showwarning("Alerta", "No esta activada produccion TAP1-WILL CALL1")
        else:
            pacifictime = datetime.now(pytz.timezone('US/pacific'))
            current_time = pacifictime.strftime("%Y-%m-%d %H:%M:%S")
            dia = pacifictime.weekday()
            FECHA = pacifictime.date()
            TIEMPO = pacifictime.time()
            if dia != 6:
                # data.append([selected_radio, selected_combobox, current_time])
                text_widget.configure(state='normal')
                text_widget.insert(tk.END, 'Orden :' + str(orden.get()) + '   ')
                text_widget.insert(tk.END,
                                   dt.strptime(current_time, "%Y-%m-%d %H:%M:%S").strftime("%a,%d %b, %Y") + '\n')
                text_widget.see(tk.END)
                text_widget.insert(tk.END, 'Store :' + str(selected_radio) + '   ')
                text_widget.see(tk.END)
                text_widget.insert(tk.END, 'Drop :' + str(selected_combobox) + '\n')
                text_widget.see(tk.END)

                text_widget.configure(state='disabled')

                fechaProm = fecha_promesa(selected_radio, selected_combobox, dia, FECHA, TIEMPO)
                if selected_radio in [4, 10, 15] and selected_combobox in ["15 EAST", "15 NORTH 2", "15 SOUTH",
                                                                           "15 WEST", "5 NORTH", "5 NORTH 2",
                                                                           "5 NORTH 3", "5 EAST", "5 WEST", "AREA SD",
                                                                           "SD AUX"
                    , "SHOP SD", "SHIPPING", "WILL CALL 1", "WILL CALL 6", "WILL CALL 7", "SHIP WC1", "SHIP WC6"]:
                    fechaProm = verifica_festivo(fechaProm, MEXICO)
                elif selected_radio in [1, 2, 3, 6, 7, 8] and selected_combobox in ["ENSENADA", "TIJUANA",
                                                                                    "PAQUETERIA TJ"]:
                    fechaProm = verifica_festivo(fechaProm, USA)
                elif selected_radio in [4, 10, 15] and selected_combobox in ["ENSENADA", "TIJUANA", "PAQUETERIA TJ"]:
                    fechaProm = verifica_festivo(fechaProm, MEXICO2)
                elif selected_radio in [1, 2, 3, 6, 7, 8] and selected_combobox in ["15 EAST", "15 NORTH 2", "15 SOUTH",
                                                                                    "15 WEST", "5 NORTH", "5 NORTH 2",
                                                                                    "5 NORTH 3", "5 EAST", "5 WEST",
                                                                                    "AREA SD", "SD AUX"
                    , "SHOP SD", "SHIPPING", "WILL CALL 1", "WILL CALL 6", "WILL CALL 7", "SHIP WC1", "SHIP WC6"]:
                    fechaProm = verifica_festivo(fechaProm, USA2)

                fechas.append(fechaProm)
                data.extend([selected])
                data.extend([str(orden.get())])
                data.extend([current_time])
                data.extend([selected_radio])
                data.extend([selected_combobox])
                data.extend([str(fechaProm)])
                df1['Part Store #'] = selected_radio
                df1['Route'] = selected_combobox
                df1['Created'] = current_time
                df1['Due Date part'] = str(fechaProm)
                text_widget.configure(state='normal')
                text_widget.insert(END,
                                   "Fecha llegada a ruta o locacion seleccionada:" + '\n' + dt.strptime(str(fechaProm),
                                                                                                        '%Y-%m-%d').strftime(
                                       "%a,%d %b, %Y") + '\n')
                text_widget.see(END)

                text_widget.configure(state='disabled')
                entry1.config(state="disabled")
                subir2()
            else:
                messagebox.showinfo("showinfo", "Domingo no es dia loaborable")
                clear_orden()

    def verifica_festivo(fechaProm, pais):

        if pais == 0 or pais == 1:
            if festivosusa.count(fechaProm.strftime('%m-%d-%Y')) or festivosusa1.count(fechaProm.strftime('%m-%d-%Y')):
                if festivosusa1.count(fechaProm.strftime('%m-%d-%Y')):
                    messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime(
                        '%m-%d-%Y') + " es 1 dia despues del festivo en USA !Consulta con Import/Export!")
                else:
                    messagebox.showerror("DIA FESTIVO", "Fecha promesa de entrega " + fechaProm.strftime(
                        '%m-%d-%Y') + " es dia festivo en USA !Consulta con Import/Export!")

            if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                if festivosmex1.count(fechaProm.strftime('%m-%d-%Y')):
                    messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime(
                        '%m-%d-%Y') + " es 1 dia despues del festivo en MEXICO !Consulta con Import/Export!")
                else:
                    messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime(
                        '%m-%d-%Y') + " es dia festivo en MEXICO !Consulta con Import/Export)")

        if pais == 2:
            if festivosmex.count(fechaProm.strftime('%m-%d-%Y')):
                if festivosmex1.count(fechaProm.strftime('%m-%d-%Y')):
                    messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime(
                        '%m-%d-%Y') + " es 1 dia despues del festivo en MEXICO !Consulta con tu encargado!")
                else:
                    messagebox.showerror("DIA FESTIVO!!!", "Fecha promesa de entrega " + fechaProm.strftime(
                        '%m-%d-%Y') + " es dia festivo en MEXICO !Consulta con tu encargado!")
        return fechaProm

    def submit_action():
        nonlocal dates
        if fechas:  # si no hay partes en la orden no calcula la Fecha promesa
            dates = largest_date(fechas)
            messagebox.showinfo("showinfo",
                                "Fecha Compromiso para el cliente es :" + dt.strptime(dates, '%Y-%m-%d').strftime(
                                    "%a,%d %b, %Y"))
            fechas.clear()
            add_button['state'] = DISABLED
            submit_button['state'] = DISABLED
            cambio_button['state'] = NORMAL
            entry1.config(state="disabled")
            clear_button['state'] = NORMAL
            data.extend(["-", "-", "-", "-", "-", "-"])
            data.extend([dates])
            subir()

    def clear_orden():

        text_widget.configure(state='normal')
        text_widget.delete(1.0, END)
        text_widget.configure(state='disabled')
        entry1.config(state="normal")
        add_button['state'] = DISABLED
        submit_button['state'] = DISABLED
        cambio_button['state'] = DISABLED
        data.clear()
        orden.set('')
        fechas.clear()
        data.extend(["-", "-", "-", "-", "-", "-", "-"])
        data.extend(["Clear"])
        subir2()

    def validation(i, text, new_text):
        return len(new_text) == 0 or len(new_text) < 10 and text.isdecimal()

    def validation2(i, text, new_text):
        return len(new_text) == 0 or (len(new_text) <= 1 and text.isalnum()) or (
                (len(new_text) > 1 and len(new_text) <= 6) and text.isdecimal())
    def form_complete(event):
        if len(orden.get()) <= 1:
            add_button['state'] = DISABLED
            submit_button['state'] = DISABLED
            cambio_button['state'] = DISABLED
            text_widget.configure(state='normal')
            text_widget.delete(1.0, END)
            text_widget.configure(state='disabled')

        else:
            add_button['state'] = NORMAL
            submit_button['state'] = NORMAL
            part_button['state'] = NORMAL

    def seleccionar():
        f = g

    def check_export():
        def close_window():
            export_window.destroy()
            main_window.attributes('-disabled', False)  # deiconify()

        """service = build('gmail', 'v1', credentials=creds)
        results = service.users().labels().list(userId='me').execute()
        send_email(service, 'tapatio.procesos@gmail.com', 'juan.a.o.delgado@gmail.com', 'Autorizacion de regresos RMA 00124', 'Esto es una prueba')    

        send_email(service, 'tapatio.procesos@gmail.com', 'jortiz@tapatioauto.com', 'Autorizacion de regresos RMA 00114', 'Esto es una prueba')    
        """

        main_window.attributes('-disabled', True)
        export_window = tk.Toplevel(main_window)
        export_window.title("Listado de partes import/export")
        export_window.geometry("350x300")
        export_window.iconbitmap("logoicon.ico")
        export_window.resizable(False, False)
        export_window.attributes('-topmost', True)

        close_button = tk.Button(export_window, text="Salir", command=close_window)
        close_button.place(x=285, y=7)
        export_window.protocol("WM_DELETE_WINDOW", close_window)

        def search_df():
            search_num = codigo_field.get()
            matches = dfpartes[
                (dfpartes['No Parte'] == str(search_num)) & (dfpartes['Destino'] == "Mexico-Estados Unidos")]
            matches2 = dfpartes[
                (dfpartes['No Parte'] == str(search_num)) & (dfpartes['Destino'] == "Estados Unidos-Mexico")]
            if (len(matches) == 0 and len(matches2) == 0):
                export_window.geometry("350x300")
                match_label.configure(text="No cruza", font=("Courier 18 bold"), fg="red")
                match_label2.configure(text="", font=("Courier 13 bold"), fg="red")
                match_label3.configure(text="", font=("Courier 13 bold"), fg="red")
                match_label4.configure(text="", font=("Courier 13 bold"), fg="red")
            else:
                export_window.geometry("1250x300")
                search_result_str = "\n".join([f"{row['Destino']} - " for _, row in matches.iterrows()])
                search_result_str3 = "\n".join([f"{row['Destino']} - " for _, row in matches2.iterrows()])
                search_result_str2 = "\n".join([f" - {row['Descripcion']}" for _, row in matches.iterrows()])
                search_result_str4 = "\n".join([f" - {row['Descripcion']}" for _, row in matches2.iterrows()])

                match_label.configure(text=search_result_str, font=("Courier 8 bold"), fg="blue")
                match_label2.configure(text=search_result_str2, font=("Courier 8 bold"), fg="black")
                match_label3.configure(text=search_result_str3, font=("Courier 8 bold"), fg="green")
                match_label4.configure(text=search_result_str4, font=("Courier 8 bold"), fg="black")
                if len(search_result_str3) == 0 and len(search_result_str) != 0:
                    match_label2.update_idletasks()
                    export_window.geometry(f"{match_label2.winfo_width() + 194}x300")

                elif len(search_result_str3) != 0 and len(search_result_str) == 0:
                    match_label4.update_idletasks()
                    export_window.geometry(f"{match_label4.winfo_width() + 194}x300")

                else:
                    match_label4.update_idletasks()
                    export_window.geometry(f"{match_label4.winfo_width() + 194}x300")

        instructions = tk.Label(export_window, text="Introduzca el codigo que esta buscando:")
        instructions.place(x=20, y=10)
        codigo_field = tk.Entry(export_window, width=5, state="normal")
        codigo_field.place(x=245, y=10)
        codigo_field.bind("<Return>", lambda event: search_df())
        codigo_field.focus()
        match_label = tk.Label(export_window, text="", justify=LEFT)
        match_label.place(x=30, y=60)
        match_label2 = tk.Label(export_window, text="", justify=LEFT)
        match_label2.place(x=180, y=60)
        match_label3 = tk.Label(export_window, text="", justify=LEFT)
        match_label3.place(x=30, y=150)
        match_label4 = tk.Label(export_window, text="", justify=LEFT)
        match_label4.place(x=180, y=150)
        export_window.mainloop()

    def submit():
        nonlocal cambio

        if cambio == 1:
            data.extend(["-", "-", "-", "-", "-", "-"])
            data.extend([dates])
            data.extend("X")
            cambio = 0
        else:
            data.extend(["-", "-", "-", "-", "-", "-"])
            data.extend([dates])
        df1['Order #'] = str(orden.get())
        df1['Due Date Order'] = dates
        df1['User'] = selected
        df1.append(data)
        cambio_button['state'] = DISABLED
        text_widget.configure(state='normal')
        text_widget.delete(1.0, END)
        text_widget.configure(state='disabled')
        entry1.config(state="normal")
        orden.set('')
        subir()

    def largest_date(dates):
        return max(dates).strftime('%Y-%m-%d')

    def fecha_promesa(selected_radio, selected_combobox, dia, FECHA, TIEMPO):
        fechaCompromiso = ''

        def tabla(tiempo, c, b):
            nonlocal fechaCompromiso
            if TIEMPO < tiempo.time():
                a = dc.get(dia)
                delt = a.get(c)
                fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours=delt[0])
            else:
                a = dc.get(dia)
                delt = a.get(b)
                fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours=delt[0])

        def tabla1(tiempo, c, b):
            nonlocal fechaCompromiso
            if TIEMPO < tiempo.time():
                fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours=0)
            else:
                fechaCompromiso = pd.to_datetime(FECHA) + timedelta(hours=0)

        St = selected_radio
        Rt = selected_combobox
        fecha1 = df.at[St, Rt]
        tiempo1 = datetime(2022, 1, 1, 12, 31, 00)
        tiempo2 = datetime(2022, 1, 1, 13, 1, 00)
        tiempo3 = datetime(2022, 1, 1, 14, 1, 00)
        tiempo4 = datetime(2022, 1, 1, 16, 1, 00)
        tiempo5 = datetime(2022, 1, 1, 17, 1,
                           00)
        tiempo6 = datetime(2022, 1, 1, 15, 1, 00)

        if fecha1 != 99:
            if dia in range(0, 6):
                if fecha1 == 1 and dia == 5:
                    tabla(tiempo2, '1.2', '1.3')
                elif fecha1 == 1:
                    tabla(tiempo4, 1, '1.1')
                if fecha1 == 2 and dia == 5:
                    tabla(tiempo2, '2.2', '2.3')
                elif fecha1 == 2:
                    tabla(tiempo4, 2, '2.1')
                if fecha1 == 3 and dia == 5:
                    tabla(tiempo2, '3.2', '3.3')
                elif fecha1 == 3:
                    tabla(tiempo4, 3, '3.1')
                if fecha1 == 4:
                    tabla(tiempo3, 4, '4.1')
                if fecha1 == 5:
                    tabla(tiempo1, 5, '5.1')
                if fecha1 == 6:
                    tabla(tiempo4, 6, '6.1')
                if fecha1 == 7:
                    tabla(tiempo3, 7, '7.1')
                if fecha1 == 8:
                    tabla(tiempo1, 8, '8.1')
                if fecha1 == 9:
                    tabla(tiempo4, 9, '9.1')
                if fecha1 == 10:
                    tabla(tiempo4, 10, '10.1')
                if fecha1 == 11:
                    tabla(tiempo4, 11, '11.1')
                if fecha1 == 12:
                    tabla(tiempo1, 12, '12.1')
                if fecha1 == 13:
                    tabla(tiempo4, 13, '13.1')
                if fecha1 == 14:
                    tabla(tiempo3, 14, '14.1')
                if fecha1 == 15:
                    tabla(tiempo1, 15, '15.1')
                if fecha1 == 16:
                    tabla(tiempo4, 16, '16.1')
                if fecha1 == 17 and dia == 5:
                    tabla(tiempo3, '17.2', '17.3')
                elif fecha1 == 17:
                    tabla(tiempo5, 17, '17.1')
                if fecha1 == 18 and dia == 5:
                    tabla(tiempo6, '18.2', '18.3')
                elif fecha1 == 18:
                    tabla(tiempo5, 18, '18.1')
                if fecha1 == 19 and dia == 5:
                    tabla(tiempo2, '19.2', '19.3')
                elif fecha1 == 19:
                    tabla(tiempo5, 19, '19.1')
                if fecha1 == 20 and dia == 5:
                    tabla(tiempo6, '20.2', '20.3')
                elif fecha1 == 20:
                    tabla(tiempo5, 20, '20.1')
            elif dia == 6:
                print("Domingo no es dia laborable")
                return ()
        elif fecha1 == 99:  # Todo lo que sea 99
            tabla1(tiempo4, 1, '1.1')
        return fechaCompromiso.date()

    def regresos():
        global input_entries, con
        input_entries.clear()
        def close_window2():
            return_window.destroy()
            main_window.attributes('-disabled', False)

        def enable_send_button():
            global input_entries
            try:
                for entry in input_entries:
                    entry_value = entry
                    print(f"Value of entry {entry}: {entry_value}")
                print(f"Media paths: {media_paths}")

                if all(entry for entry in input_entries) and media_paths:
                    send_button.config(state=tk.NORMAL)
                else:
                    send_button.config(state=tk.DISABLED)
            except Exception as e:
                print(f"An error occurred: {str(e)}")

        def open_file_dialog():
            allowed_extensions = [".jpg", ".jpeg", ".png", ".gif", ".mp4", ".avi", ".mov"]
            file_paths = filedialog.askopenfilenames(
                title="Seleccionar archivos multimedia",
                filetypes=[("Archivos multimedia", allowed_extensions)]
            )

            for file_path in file_paths:
                mimetype, _ = mimetypes.guess_type(file_path)
                if mimetype and (mimetype.startswith('image') or mimetype.startswith('video')):
                    media_paths.append(file_path)

            return_window.lift()
            enable_send_button()

        def verifica_garantia():
            fecha_venta = fecha_venta_cal.get_date()
            days_difference = (today - fecha_venta).days
            var1 = tk.IntVar()
            if tipo_parte_menu.get() in Garantia90:
                if days_difference < 91:
                    stock_entry.config(state="normal")
                else:
                    messagebox.showinfo("Garantía No Aplica",
                                        "La garantía no aplica porque han pasado más de 90 días desde la fecha de venta.")
                    tipo_parte_menu.set("")
                    tipo_parte_menu.config(state="disabled")
                    return_window.lift()
            elif tipo_parte_menu.get() in Nogarantia:
                if days_difference > 30:
                    messagebox.showinfo("Garantía No Aplica",
                                        "La garantía no aplica porque han pasado más de 30 días desde la fecha de venta.")
                    tipo_parte_menu.set("")
                    tipo_parte_menu.config(state="disabled")
                else:
                    messagebox.showinfo("Parte no tiene garantia", "Este tipo de parte no tiene garantia")
                    garantia_opcion = tk.Radiobutton(return_window, text="Garantia", variable=var1, value=1,
                                                     command=garantia_parte).place(x=280, y=100)
                    return_window.lift()
            else:
                if days_difference < 31:
                    stock_entry.config(state="normal")
                else:
                    messagebox.showinfo("Garantía No Aplica",
                                        "La garantía no aplica porque han pasado más de 30 días desde la fecha de venta.")
                    tipo_parte_menu.delete(0, tk.END)
                    tipo_parte_menu.config(state="disabled")
                    return_window.lift()

        def guardar_datos():
            global con
            abrir_conexion()
            invoice = invoice_entry.get()
            fecha_venta = fecha_venta_cal.get_date()
            tipo_parte = input_entries[3].get()
            num_stock = stock_entry.get().upper()
            motivo_regreso = motivo_regreso_menu.get()
            vendedor = input_entries[2].get()
            dia_req = datetime.today()
            descripcion_motivo = descripcion_motivo_text.get("1.0", "end")
            autorizado = "Pendiente"
            fecha_autorizacion = None
            persona_autorizacion = None
            num_orden_regreso = None
            estatus_regreso = "PendingAUT"

            fecha_venta = fecha_venta.strftime('%Y/%m/%d')
            dia_req = dia_req.strftime('%Y/%m/%d')
            temporal = folder["id"]
            pictures = f"https://drive.google.com/drive/folders/{folder_id}/{temporal}"

            try:
                cursor = con.cursor()
                cursor.execute("""
                INSERT INTO regresos (Invoice, Fecha_Venta, Tipo_Parte, Num_Stock, Motivo_Regreso, Descripcion_Motivo,
                                      Autorizado, Fecha_Autorizacion, Persona_Autorizacion, Num_Orden_Regreso, Estatus_Regreso, User_App,Seller,Fecha_Req,Fotos)
                VALUES (:1, TO_DATE(:2, 'YYYY/MM/DD'), :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13,TO_DATE(:14, 'YYYY/MM/DD'), :15)
                """, (invoice, fecha_venta, tipo_parte, num_stock, motivo_regreso, descripcion_motivo,
                      autorizado, fecha_autorizacion, persona_autorizacion, num_orden_regreso, estatus_regreso,
                      selected, vendedor, dia_req, pictures))
                con.commit()

            except oracledb.Error as error:
                print(f"Error de Oracle: {error}")
            finally:
                con.close()

        def send_data():
            global folder
            invoice = invoice_entry.get()
            fecha_venta = fecha_venta_cal.get_date()
            tipo_parte = tipo_parte_menu.get()
            motivo_regreso = motivo_regreso_var.get()
            vendedor = input_entries[2].get()
            descripcion_motivo = descripcion_motivo_text.get("1.0", tk.END)
            stock = stock_entry.get()

            fecha_venta_str = fecha_venta_cal.get_date().strftime("%d/%m/%Y")
            fecha_requerimiento = today.strftime("%d/%m/%Y")
            new_row = [
                selected,
                fecha_requerimiento,
                input_entries[0].get(),
                fecha_venta_str,
                tipo_parte_menu.get(),
                input_entries[3].get(),
                motivo_regreso_menu.get(),
                descripcion_motivo_text.get("1.0", "end"),
                input_entries[2].get(),

            ]

            folder_name = "invoice_" + invoice
            try:
                folder_metadata = {
                    'name': folder_name,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'parents': [folder_id]
                }
                folder = service.files().create(body=folder_metadata, fields='id').execute()
            except HttpError as e:
                print(f"Error HTTP: {e}")
            except Exception as e:
                print(f"Error: {e}")
            print(f'Carpeta creada con ID: {folder["id"]}')
            temporal = folder["id"]
            pictures = f"https://drive.google.com/drive/folders/{folder_id}/{temporal}"
            print(pictures)

            for index, media_path in enumerate(media_paths):
                if not invoice:
                    messagebox.showerror("Error", "Debes ingresar un número de factura.")
                    return
                nuevo_nombre = f"invoice_{invoice}_{tipo_parte}_{index + 1}"
                print("SUBIENDO ARCHIVO # ", index + 1)

                file_metadata = {
                    'name': nuevo_nombre,
                    'parents': [folder["id"]]
                }

                try:
                    media = MediaFileUpload(media_path, mimetype=mimetypes.guess_type(media_path)[0])
                    service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id'
                    ).execute()
                except HttpError as e:
                    print(f"Error HTTP: {e}")
                except Exception as e:
                    print(f"Error: {e}")

            listado.append_row(new_row)
            guardar_datos()
            send_email(tipo_parte, invoice, fecha_venta, motivo_regreso, descripcion_motivo, stock, vendedor)

        def send_email(tipo_parte, invoice, fecha_venta, motivo_regreso, descripcion_motivo, stock, vendedor):
            message = create_message(tipo_parte, invoice, fecha_venta, motivo_regreso, descripcion_motivo, stock,
                                     vendedor)
            max_retries = 3
            retries = 0
            while retries < max_retries:
                try:
                    message = (services.users().messages().send(userId='me', body=message).execute())
                    print('Mensaje enviado. ID del mensaje: {}'.format(message['id']))

                    break
                except Exception as e:
                    messagebox.showwarning("Send error",
                                           'Error al enviar el mensaje: {}...Reintentando en 5 segundos...'.format(e))
                    time.sleep(5)
                    retries += 1

            else:
                messagebox.showerror("Error de Envío",
                                     "No se pudo enviar el mensaje después de {} intentos.".format(max_retries))
                return

            messagebox.showinfo("Envio de email exitoso", "La informacion y las fotos se enviaron correctamente.")
            close_window2()
            listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button)

        def create_message(tipo_parte, invoice, fecha_venta, motivo_regreso, descripcion_motivo, stock, vendedor):

            cadena = tipo_parte_menu.get()
            Code = cadena[:3]
            invoice = invoice
            temporal = folder["id"]
            message = MIMEMultipart()
            message[
                'to'] = 'help@tapatioauto.com'
            message['cc'] = 'paraujo@tapatioauto.com'
            message['bcc'] = 'juan.a.o.delgado@gmail.com'
            message['subject'] = 'Solicitud de autorizacion de regreso de partes'

            body = f"""\
            Hola,
            Por favor revisar la siguiente informacion para la 
            autorizacion del regreso de la parte {cadena}\n
            Invoice: {invoice}
            Fecha de Venta: {fecha_venta}
            Vendedor: {vendedor}
            Tipo de parte: {Code}
            # de Stock: {stock}
            Motivo del Regreso: {motivo_regreso_menu.get()}
            Descripcion del motivo: {descripcion_motivo}

            Puedes acceder a las fotos y videos en el siguiente enlace: 
            
            (https://drive.google.com/drive/folders/{folder_id}/{temporal})

            Adjunto encontrarás las fotos.

            Saludos,
            {selected}
            """
            message.attach(MIMEText(body, 'plain'))
            for indexx, media_path in enumerate(media_paths):

                mimetype, _ = mimetypes.guess_type(media_path)
                nuevo_nombre = f"invoice_{invoice}_{tipo_parte}_{indexx + 1}"
                if mimetype and mimetype.startswith('image'):
                    # Es una foto
                    image = MIMEImage(open(media_path, 'rb').read(), name=nuevo_nombre)
                    message.attach(image)

            # Enviamos el mensaje
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
            return {'raw': raw_message}

        def enable_tipo(event):
            if fecha_venta_cal.get() != "":
                tipo_parte_menu.config(state="normal")
            else:
                tipo_parte_menu.config(state="disabled")

        def enable_calendar(event):
            fecha_venta_cal.config(state="normal")

        def garantia_parte():
            stock_entry.config(state="normal")

        def enable_motivo(event):
            motivo_regreso_menu.config(state="readonly")

        def enable_descripcion(event):
            descripcion_motivo_text.config(state="normal")

        def enable_file_button(event):
            if descripcion_motivo_text.get("1.0", "end-1c"):
                file_button.config(state="normal")
            else:
                file_button.config(state="disabled")

        def enable_file(event):
            descripcion_motivo_text.config(state="normal")

            enable_file_button(None)

        def is_number(char):
            try:
                int(char)
                return True
            except ValueError:
                return False

        def show_listbox(entry):
            global lista
            lista = tk.Listbox(return_window, height=3, width=30)
            update(tipo_partes)
            llenar()

            def select_item():
                selected = lista.get(lista.curselection())
                entry.delete(0, tk.END)
                entry.insert(0, selected)
                lista.destroy()
                verifica_garantia()

            lista.bind("<Double-Button-1>", lambda e: select_item())
            lista.place(x=entry.winfo_x(), y=entry.winfo_y() + entry.winfo_height())

        def update(data):
            global lista
            lista.delete(0, END)

            for item in data:
                lista.insert(END, item)

        def llenar():
            global lista
            tipo_parte_menu.delete(0, END)
            tipo_parte_menu.insert(0, lista.get(ANCHOR))

        def revisar(event):
            typed = tipo_parte_menu.get()

            if typed == '':
                data = tipo_partes
            else:
                data = []
                for item in tipo_partes:
                    if typed.lower() in item.lower():
                        data.append(item)

            update(data)

        def toggle_combobox_state():
            if variable.get() == 1:
                vendedor_menu.config(state="readonly")
            else:
                vendedor_menu.set("")
                vendedor_menu.config(state="disabled")

        main_window.attributes('-disabled', True)
        return_window = tk.Toplevel(main_window)
        return_window.title("Cuestionario para regresos")
        return_window.geometry("400x450")
        return_window.iconbitmap("logoicon.ico")
        return_window.resizable(False, False)
        return_window.attributes('-topmost', True)
        input_labels = ["Invoice:", "Fecha de Venta:", "Vendedor:", "Tipo de parte:", "# de Stock:",
                        "Motivo del Regreso:", "Descripcion del motivo:"]
        # input_entries = []
        variable = 0
        today = datetime.now().date()
        for label_text in input_labels:
            label = tk.Label(return_window, text=label_text)
            label.pack()
            if label_text == "Invoice:":
                Invoice = tk.StringVar(return_window)
                invoice_entry = tk.Entry(return_window, textvariable=Invoice, width=10, validate="key",
                                         validatecommand=(return_window.register(validation), "%d", "%S", "%P"))
                invoice_entry.pack()
                input_entries.append(invoice_entry)
                invoice_entry.focus()
            elif label_text == "Tipo de parte:":
                tipo_parte_menu = tk.Entry(return_window, width=30, state="disabled")
                tipo_parte_menu.pack()
                tipo_parte_menu.bind("<FocusIn>", lambda e: show_listbox(tipo_parte_menu))
                tipo_parte_menu.bind("<KeyRelease>", revisar)
                input_entries.append(tipo_parte_menu)
            elif label_text == "Vendedor:":
                variable = tk.IntVar()
                sel_vendedor = tk.Radiobutton(return_window, variable=variable, value=1, command=toggle_combobox_state)
                vendedor_menu = ttk.Combobox(return_window, values=list(users.keys()))
                vendedor_menu.current(list(users.keys()).index(selected))
                vendedor_menu.pack()
                vendedor_menu.config(state="disabled")
                sel_vendedor.place(x=105, y=102)
                vendedor_menu.bind("<<ComboboxSelected>>", enable_descripcion)
                input_entries.append(vendedor_menu)
            elif label_text == "Fecha de Venta:":
                fecha_venta_cal = DateEntry(return_window, date_pattern='dd/mm/yyyy')
                fecha_venta_cal.pack()
                fecha_venta_cal.bind("<<DateEntrySelected>>", enable_tipo)
                fecha_venta_cal.config(state="disabled")
                invoice_entry.bind("<FocusIn>", enable_calendar)
                input_entries.append(fecha_venta_cal)
                # print("hola")
            elif label_text == "# de Stock:":
                StockN = tk.StringVar(return_window)
                stock_entry = tk.Entry(return_window, textvariable=StockN, width=10, validate="key",
                                       validatecommand=(return_window.register(validation2), "%d", "%S", "%P"))
                stock_entry.pack()
                input_entries.append(stock_entry)
                stock_entry.config(state="disabled")
                stock_entry.bind("<FocusIn>", enable_motivo)
            elif label_text == "Motivo del Regreso:":
                motivo_regreso_var = tk.StringVar(return_window)
                motivo_regreso_options = [
                    "1.- Not Needed",
                    "2.- Bad Quality",
                    "3.- Damage/Shipping",
                    "4.- Damage/Defected",
                    "5.- Damage at Yard",
                    "6.- Wrong Part",
                    "7.- Wrong part Interchange",
                    "8.- Wrong Part/Bad Inventory",
                    "9.- Other",
                    "10.- Invalid",
                    "11.- Missing Part",
                    "12.- Wrong Part/Customer",
                    "13.- No production"
                ]
                motivo_regreso_menu = ttk.Combobox(return_window, state="readonly", values=motivo_regreso_options)
                motivo_regreso_menu.pack()
                motivo_regreso_menu.config(state="disabled")
                motivo_regreso_menu.bind("<<ComboboxSelected>>", enable_descripcion)
                input_entries.append(motivo_regreso_menu)
            elif label_text == "Descripcion del motivo:":
                descripcion_motivo_text = tk.Text(return_window, height=7, width=40)
                descripcion_motivo_text.pack()
                descripcion_motivo_text.config(state="disabled")
                descripcion_motivo_text.bind("<FocusIn>", enable_file)
                descripcion_motivo_text.bind("<KeyRelease>", enable_file)
                input_entries.append(descripcion_motivo_text.get("1.0", tk.END))
                # print(input_entries[5])
        file_button = tk.Button(return_window, text="Cargar Fotos", command=open_file_dialog)
        file_button.pack()
        file_button.config(state="disabled")
        send_button = tk.Button(return_window, text="Enviar", command=send_data, state=tk.DISABLED)
        send_button.pack()
        cancel_button = tk.Button(return_window, text="Salir", command=close_window2, state=tk.NORMAL)
        cancel_button.place(x=350, y=410)
        return_window.protocol("WM_DELETE_WINDOW", close_window2)

        return_window.mainloop()

    def consultar_rt_dt(event):
        global con
        if QtyRtnaut:
            abrir_conexion()
            try:
                cursor = con.cursor()
                if selected in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO", "EMMANUEL LOPEZ"]:
                    cursor.execute(
                        "SELECT * FROM regresos WHERE ESTATUS_REGRESO = :1 AND AUTORIZADO = :2 AND NUM_ORDEN_REGRESO IS NULL",
                        ('PendingDT', 'Yes'))
                else:
                    cursor.execute(
                        "SELECT * FROM regresos WHERE SELLER = :1 AND ESTATUS_REGRESO = :2 AND AUTORIZADO = :3 AND NUM_ORDEN_REGRESO IS NULL",
                        (selected, 'PendingDT', 'Yes'))

                resultados = cursor.fetchall()

                if resultados:

                    DespliegueDB.main_function(selected, main_window, resultados, 0, date.today())
                    listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button)
                else:
                    messagebox.showinfo("Sin Resultados",
                                        "No se encontraron DT pendientes de generar para este vendedor.")

            except oracledb.Error as error:
                print(f"Error de Oracle: {error}")

    def consultar_rt_pen(event):
        global con
        if QtyRtnPending:
            abrir_conexion()
            try:
                cursor = con.cursor()
                if selected in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO", "EMMANUEL LOPEZ"]:
                    cursor.execute(
                        "SELECT * FROM regresos WHERE ESTATUS_REGRESO = :1 AND NUM_ORDEN_REGRESO IS NULL",
                        ('PendingAUT',))
                else:
                    cursor.execute(
                        "SELECT * FROM regresos WHERE SELLER = :1 AND ESTATUS_REGRESO = :2 AND NUM_ORDEN_REGRESO IS NULL",
                        (selected, 'PendingAUT'))

                resultados = cursor.fetchall()

                if resultados:
                    DespliegueDB.main_function(selected, main_window, resultados, 1, date.today())
                    listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button)
                else:
                    messagebox.showinfo("Sin Resultados",
                                        "No se encontraron Devoluciones pendientes de revisar para este vendedor.")

            except oracledb.Error as error:
                print(f"Error de Oracle: {error}")

    def consultar_rt_reje(event):
        global con
        if QtyRtnNG:
            abrir_conexion()
            try:
                cursor = con.cursor()
                if selected in ["MARIANA MENDOZA", "JUAN ORTIZ", "PEDRO JUVERA", "MANUEL RAZO", "EMMANUEL LOPEZ"]:
                    cursor.execute(
                        "SELECT * FROM regresos WHERE AUTORIZADO = : 1 AND ESTATUS_REGRESO = :2 AND NUM_ORDEN_REGRESO IS NULL",
                        ('No', 'DoneNG'))
                else:
                    cursor.execute(
                        "SELECT * FROM regresos WHERE SELLER = :1 AND AUTORIZADO = :2 AND ESTATUS_REGRESO = :3 AND NUM_ORDEN_REGRESO IS NULL",
                        (selected, 'No', 'DoneNG'))
                resultados = cursor.fetchall()

                if resultados:
                    DespliegueDB.main_function(selected, main_window, resultados, 2, date.today())
                    listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button)
                else:
                    messagebox.showinfo("Sin Resultados",
                                        "No se encontraron Devoluciones no autorizadas para este vendedor.")

            except oracledb.Error as error:
                print(f"Error de Oracle: {error}")

    main_window = tk.Tk()
    main_window.title("Fecha compromiso TAPATIO")
    main_window.geometry("375x285")  # x,y)
    main_window.iconbitmap("logoicon.ico")
    main_window.resizable(False, False)
    image = Image.open("logo-new.png")
    image = image.resize((70, 30), Image.LANCZOS)
    image = ImageTk.PhotoImage(image)
    label_image = tk.Label(image=image).place(x=0, y=1)

    label = tk.Label(main_window, text="", font=("times", 8))
    label.place(x=80, y=2)
    if super_user == 1:
        label.config(text=selected + " (SUPER USER)")
    else:
        label.config(text=selected)
    main_window.bind("<KeyRelease>", form_complete)
    clock = tk.Label(main_window, justify=tk.RIGHT, font=("times", 9, "bold"), fg="blue")
    clock.place(x=90, y=16)

    def change_promise_day():
        nonlocal cambio
        calendar_window = tk.Tk()
        calendar_window.title("Login")
        calendar_window.geometry("300x300")
        calendar_window.iconbitmap("logoicon.ico")
        calendar_window.resizable(False, False)
        calendar_window.attributes('-topmost', True)

        def grad_date():
            nonlocal dates, cambio
            c = cal.get_date()
            d = dt.strptime(c, '%m/%d/%y')
            e = dt.strptime(dates, '%Y-%m-%d')
            if e < d:
                messagebox.showinfo("showinfo",
                                    "La nueva Fecha Compromiso para el cliente cambio de :" + '\n' + dt.strptime(dates,
                                                                                                                 '%Y-%m-%d').strftime(
                                        "%a,%d %b, %Y") +
                                    "  a  " + dt.strptime(cal.get_date(), '%m/%d/%y').strftime("%a,%d %b, %Y"))

                dates = dt.strptime(cal.get_date(), '%m/%d/%y').strftime("%Y-%m-%d")
                df1['DueDate change'] = '*'
                data.extend(["-", "-", "-", "-", "-", "-"])
                data.extend([dates])
                data.extend("X")
                cambio = 0
                calendar_window.destroy()
                subir()

        cal = Calendar(calendar_window, selectmode='day',
                       year=current_time.year, month=current_time.month,
                       day=current_time.day)

        cal.pack(pady=20)
        Button(calendar_window, text="Cambiar Fecha Compromiso", command=grad_date).pack(pady=5)
        date = Label(main_window, text="")
        date.place(x=200, y=450)
        cambio = 1

    def subir():
        hoja.append_row(data)
        data.clear()

    def close_all():
        global detener_thread, detener_thread2, con

        try:
            con.close()

        except oracledb.Error as error:
            print("Conexion a DB terminada")
            pass

        detener_thread = True
        detener_thread2 = True
        thread2.join()
        thread3.join()
        main_window.destroy()
        login()

    def subir2():
        hoja.append_row(data)
        data.clear()

    def tick():
        time_string = time.strftime("%H:%M:%S")
        clock.config(text=time_string)
        clock.after(200, tick)

    tick()
    var = tk.IntVar()
    var.set(1)
    radio1 = tk.Radiobutton(main_window, text="TAP 1", variable=var, value=1)
    radio2 = tk.Radiobutton(main_window, text="TAP 2", variable=var, value=2)
    radio3 = tk.Radiobutton(main_window, text="TAP 4 & 14", variable=var, value=4)
    radio4 = tk.Radiobutton(main_window, text="TAP 6", variable=var, value=6)
    radio5 = tk.Radiobutton(main_window, text="TAP 7 & 8", variable=var, value=7)
    radio6 = tk.Radiobutton(main_window, text="TAP 10", variable=var, value=10)
    radio7 = tk.Radiobutton(main_window, text="TAP 15", variable=var, value=15)
    radio8 = tk.Radiobutton(main_window, text="TAP 20", variable=var, value=20)

    radio1.place(x=0, y=70)
    radio2.place(x=0, y=90)
    radio3.place(x=60, y=70)
    radio4.place(x=60, y=90)
    radio5.place(x=150, y=70)
    radio6.place(x=150, y=90)
    radio7.place(x=230, y=70)
    radio8.place(x=230, y=90)
    radio8['state'] = tk.DISABLED

    part_button = tk.Button(main_window, text="Import/Export", command=check_export,
                            state=tk.NORMAL)
    part_button.place(x=4, y=255)

    if super_user:
        var2 = tk.IntVar()
        check = tk.Checkbutton(main_window, text="Hora corte (manual)", font=("Courier 10 bold"), fg="black",
                               variable=var2, onvalue=1, offvalue=0, command=seleccionar)
        check.place(x=0, y=450)

    global entry1

    name = Label(main_window, font=("Courier 10 bold"), text="Job #:").place(x=4, y=40)
    orden = tk.StringVar(main_window)
    entry1 = tk.Entry(main_window, textvariable=orden, width=10, validate="key",
                      validatecommand=(main_window.register(validation), "%d", "%S", "%P"))
    entry1.place(x=50, y=40)
    combobox = ttk.Combobox(main_window, state="readonly", values=[
        "15 EAST", "15 NORTH 2", "15 SOUTH", "15 WEST", "5 NORTH", "5 NORTH 2", "5 NORTH 3", "5 EAST", "5 WEST",
        "AREA SD", "SD AUX"
        , "SHOP SD", "ENSENADA", "TIJUANA", "EBAY TJ", "SHIPPING", "WILL CALL 1", "WILL CALL 6", "WILL CALL 7",
        "PAQUETERIA TJ"
        , "SHIP WC1", "SHIP WC6"])
    combobox.place(x=5, y=120)
    combobox.set("15 EAST")

    tk.Label(main_window, text='Ver.13 Nov/2023', font=("Courier 7 bold"), fg="black").place(x=290, y=2)
    Cambio = tk.Label(main_window, font=("times", 10, "bold"), fg="black", text="Tipo_Cambio").place(x=151, y=16)
    T_Cambio = tk.Label(main_window, font=("times", 13, "bold"), fg="red", text=f"$ {Tipo_Cambio}")
    T_Cambio.place(x=160, y=37)
    add_button = tk.Button(main_window, text="Add", command=add_action, state=tk.DISABLED)
    add_button.place(x=150, y=117)
    main_window.bind('<Return>', lambda event=None: add_button.invoke())
    submit_button = tk.Button(main_window, text="Calcular", command=submit_action,
                              state=tk.DISABLED)
    submit_button.place(x=187, y=117)

    cambio_button = tk.Button(main_window, text="Cambiar", command=change_promise_day, state=tk.DISABLED)
    cambio_button.place(x=245, y=117)

    text_widget = tk.Text(main_window, state='disabled', height=6, width=37)
    text_widget.place(x=4, y=150)

    tk.Label(main_window, text=selected_user).place(x=430, y=450)
    clear_button = tk.Button(main_window, text="Clear", command=clear_orden, state=tk.NORMAL)
    clear_button.place(x=117, y=37)

    tk.Label(main_window, text='Estatus regresos', font=("Courier 10 bold"), fg="black").place(x=240, y=16)

    RtPndaut_button = tk.Label(main_window, font=("times", 17, "bold"), fg="green")
    RtPndaut_button.place(x=260, y=35)

    RtPndaut_button.bind("<Button-1>", consultar_rt_dt)

    RtPnd_button = tk.Label(main_window, font=("times", 17, "bold"), fg="orange")
    RtPnd_button.place(x=290, y=35)
    RtPnd_button.bind("<Button-1>", consultar_rt_pen)

    Rtng_button = tk.Label(main_window, font=("times", 17, "bold"), fg="red")
    Rtng_button.place(x=330, y=35)

    Rtng_button.bind("<Button-1>", consultar_rt_reje)

    tk.Button(main_window, text="Salir", command=close_all).place(x=270, y=255)
    Returns = tk.Button(main_window, text="regresos", fg="red", command=regresos)
    Returns.place(x=95, y=255)

    listado_regresos_status(listado, RtPnd_button, RtPndaut_button, Rtng_button)
    data = []
    fechas = []
    dates = ''
    cambio = 0

    thread2 = threading.Thread(target=actualizacion)
    thread3 = threading.Thread(target=conexion)
    thread2.daemon = True
    thread3.daemon = True
    thread2.start()
    thread3.start()

    main_window.attributes('-topmost', True)
    main_window.protocol("WM_DELETE_WINDOW", close_all)
    main_window.mainloop()


def login():
    def show_listbox(entry):
        global lista
        lista = tk.Listbox(login_window, height=3, width=30)
        update(users.keys(), lista)
        llenar(lista)

        def select_item():
            selected = lista.get(lista.curselection())
            entry.delete(0, tk.END)
            entry.insert(0, selected)
            lista.destroy()
            enter_button.config(state='normal')

        lista.bind("<Double-Button-1>", lambda e: select_item())
        lista.place(x=entry.winfo_x(), y=entry.winfo_y() + entry.winfo_height())

    def update(data, lista):
        lista.delete(0, END)
        for item in data:
            lista.insert(END, item)

    def llenar(lista):

        combobox.delete(0, END)
        combobox.insert(0, lista.get(ANCHOR))

    def revisar(event):
        global lista
        typed = combobox.get()

        if typed == '':
            data = users.keys()
        else:
            data = []
            for item in users.keys():
                if typed.lower() in item.lower():
                    data.append(item)

        update(data, lista)

    def close_login():
        login_window.destroy()

    def enter(select):
        global selected
        selected = select
        super_user = 0
        password = simpledialog.askstring("Password", "Enter the password for '" + selected + "':", show='*')
        if password is not None:
            if password == users[selected]:
                if selected in ["MANUEL RAZO", "EMMANUEL LOPEZ", "KARLA CRUZ", "MIGUEL CERVANTES", "JUAN ORTIZ"]:
                    super_user = 1
                hoja = ss.worksheet(selected)
                listado = ss.worksheet("LISTADO REGRESOS")
                login_window.destroy()
                main_program(super_user, hoja, listado)
            else:
                messagebox.showerror("Error", "Incorrect password")

    login_window = tk.Tk()
    login_window.title("Login")
    login_window.geometry("300x150")
    login_window.iconbitmap("logoicon.ico")
    login_window.resizable(False, False)

    imag = Image.open("logo-new.png")
    imag1 = imag.resize((130, 80), Image.LANCZOS)
    imag1 = ImageTk.PhotoImage(imag1)
    label_image2 = tk.Label(image=imag1).place(x=0, y=40)

    image = Image.open("motor.png")
    image1 = image.resize((150, 110), Image.LANCZOS)
    image1 = ImageTk.PhotoImage(image1)
    label_image1 = tk.Label(image=image1).place(x=160, y=20)

    combobox = tk.Entry(login_window, width=30)
    combobox.pack()
    combobox.bind("<FocusIn>", lambda e: show_listbox(combobox))
    combobox.bind("<KeyRelease>", revisar)
    enter_button = tk.Button(login_window, text="Enter", command=lambda: enter(combobox.get()))
    enter_button.pack()
    enter_button.config(state='disabled')
    login_window.bind('<Return>', lambda event=None: enter_button.invoke())
    login_window.attributes('-topmost', True)
    login_window.protocol("WM_DELETE_WINDOW", close_login)
    login_window.mainloop()


selected_user = ""
login()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
