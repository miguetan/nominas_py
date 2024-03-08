import os
from PyPDF2 import PdfReader, PdfWriter
import re
import datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time


archivo_pdf = "nominas.pdf"

fecha_actual = datetime.datetime.now()
nombre_log = "Resultado " + fecha_actual.strftime("%Y_%m_%d_%H_%M_%S") + ".txt"


def log(mensaje):
# Crear el archivo de log
    with open(nombre_log, "a") as log_file:
        print(mensaje)
        log_file.write(f"{fecha_actual.strftime('%H:%M:%S')} - {mensaje}\n")


def buscar_dni(pdf_path):
    # Creamos un objeto PdfReader para el archivo PDF
    with open(pdf_path, "rb") as archivo:
        pdf_reader = PdfReader(archivo)

        # Iteramos sobre cada página del PDF
        for pagina in pdf_reader.pages:
            # Buscamos el patrón de un campo DNI en el texto de la página
            dni_pattern = r"\b\d{8}[A-HJ-NP-TV-Z]\b"
            match = re.search(dni_pattern, pagina.extract_text())

            # Si se encuentra un campo DNI, lo retornamos
            if match:
                return match.group()

    # Si no se encuentra ningún campo DNI, retornamos None
    return None


def buscar_email_por_dni(dni):
    # Cargamos el archivo Excel
    df = pd.read_excel('emails.xls')

    # Buscamos el DNI en la primera columna del DataFrame
    resultado = df[df.iloc[:, 0] == dni]

    # Si encontramos el DNI, devolvemos el email de la segunda columna
    if len(resultado) > 0:
        return resultado.iloc[0, 1]
    else:
        log(f"*****ERROR!! No se encuentra el DNI {dni} en emails.xls")
        return False

def enviar_email(destinatario, asunto, archivo_adjunto):
    log("Abriendo el archivo de configuración config.txt")
    # Leer el archivo de configuración
    with open("config.txt", "r") as config_file:
        config_data = config_file.read()

    # Obtener los valores de configuración
    remitente = re.search(r"remitente=(.*)", config_data).group(1)
    password = re.search(r"password=(.*)", config_data).group(1)
    servidor_smtp = re.search(r"servidor=(.*)", config_data).group(1)
    puerto_smtp = int(re.search(r"puerto=(.*)", config_data).group(1))

    log("Configuración leída correctamente")

    # Crear el objeto MIMEMultipart para el correo electrónico
    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto

    # Adjuntar el archivo al correo electrónico
    log("Abriendo el archivo " + archivo_adjunto + " para ser adjuntado al correo electrónico")
    adjunto = open(archivo_adjunto, "rb")
    parte_adjunta = MIMEBase("application", "octet-stream")
    parte_adjunta.set_payload((adjunto).read())
    encoders.encode_base64(parte_adjunta)
    parte_adjunta.add_header("Content-Disposition", "attachment", filename=archivo_adjunto)
    log("Adjuntando el archivo " + archivo_adjunto + " al correo electrónico")
    mensaje.attach(parte_adjunta)

    # Conectar al servidor SMTP y enviar el correo electrónico
    try:
        log("Conectando al servidor SMTP " + servidor_smtp + " en el puerto " + str(puerto_smtp) )
        servidor = smtplib.SMTP_SSL(servidor_smtp, str(puerto_smtp))
        log("Identificandose con remitente " + remitente + " y contraseña " + password)
        servidor.login(remitente, password)
        log("Conexión al servidor SMTP establecida correctamente")
        log("Enviando correo electrónico a " + destinatario)
        servidor.sendmail(remitente, destinatario, mensaje.as_string())
        servidor.quit()
        log("Correo electrónico enviado correctamente a " + destinatario  + "\n\n\n")
        return True
    except Exception as e:
        log("*******Error al enviar el correo electrónico a " + destinatario + "\n\n\n")
        print("Error al enviar el correo electrónico:", str(e))
        return False



# Comprobamos si el archivo existe
if os.path.isfile(archivo_pdf):
    # Abrimos el archivo PDF
    with open(archivo_pdf, "rb") as archivo:
        # Creamos un objeto PdfReader
        pdf_reader = PdfReader(archivo)

        # Obtenemos el número total de páginas
        num_paginas = len(pdf_reader.pages)

        # Iteramos sobre cada página del PDF
        for pagina in range(num_paginas):
            # Creamos un objeto PdfWriter para cada página
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[pagina])


            # Guardamos la página en un archivo separado
            nombre_archivo = f"nomina_{pagina + 1}.pdf"

            with open(nombre_archivo, "wb") as archivo_salida:
                pdf_writer.write(archivo_salida)
            
            # Buscamos el DNI en la página actual
            dni = buscar_dni(nombre_archivo)

            # Si se encuentra un DNI, lo imprimimos
            if dni:
                log(f"Encontrado DNI {dni} en la nomina_{pagina + 1}.pdf")
                email = buscar_email_por_dni(dni)
                if email:
                    log(f"Email asociado: {email}")
                    enviar_email(email,"Adjuntamos nómina ", nombre_archivo)
                    time.sleep(5) # Esperamos 5 segundos para no saturar el servidor de correo

            else:
                log(f"¡¡ERROR!! No se encuentra ningun DNI en nomina_{pagina + 1}.pdf")
            
            # Eliminamos el archivo temporal
            os.remove(nombre_archivo)



