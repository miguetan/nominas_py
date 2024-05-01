import os
import shutil
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
# Borramos la carpeta "errores" con todo su contenido
if os.path.exists("errores"):
    shutil.rmtree("errores")


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
            dni_pattern = r"\b\d{8}[A-HJ-NP-TV-Z]\b|\b[A-HJ-NP-TV-Z]\d{7}[A-HJ-NP-TV-Z]\b"
            match = re.search(dni_pattern, pagina.extract_text())

            # Si se encuentra un campo DNI, lo retornamos
            if match:
                return match.group()

    # Si no se encuentra ningún campo DNI, retornamos None
    return None

def buscar_email_por_dni(dni,hoja=0):
    filename = 'emails.xlsx'
    empresas = pd.ExcelFile('emails.xlsx').sheet_names

    # Leer el archivo de emails
    df = pd.read_excel(filename, sheet_name=hoja)


    # Buscar el DNI en la primera columna
    resultado = df[df.iloc[:, 0] == dni]

    # Obtener el nombre de la empresa
    empresa = empresas[hoja].lower()

    # Si encontramos el DNI, devolvemos el email de la segunda columna
    if len(resultado) > 0:
        log(f"Encontrado DNI {dni} en pestaña {empresa} del archivo emails.xls")
        return resultado.iloc[0, 1], empresa.lower()
    else:
        if hoja < len(empresas) - 1:
            return buscar_email_por_dni(dni, hoja + 1)
        log(f"*****ERROR!! No se encuentra el DNI {dni} en emails.xls")
        return False, False

def enviar_email(destinatario, asunto, archivo_adjunto, empresa):
    # Leer el archivo de configuración
    with open("config.txt", "r") as config_file:
        config_data = config_file.read()

    # Obtener los valores de configuración
    try:
        remitente = re.search(r"remitente_" + empresa +"=(.*)", config_data).group(1)
        password = re.search(r"password_" + empresa +"=(.*)", config_data).group(1)
        servidor_smtp = re.search(r"servidor_" + empresa +"=(.*)", config_data).group(1)
        puerto_smtp = int(re.search(r"puerto_" + empresa +"=(.*)", config_data).group(1))
        log("Remitente: " + remitente)
    except Exception as e:
        log("*****Error al leer la configuración del archivo config.txt")
        return False

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


# Comprobamos si el archivo de nominas existe
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
                email, empresa = buscar_email_por_dni(dni)
                if email and empresa:
                    log(f"Email asociado: {email} Empresa asociada: {empresa}")

                    enviar_email(email,"Adjuntamos nómina ", nombre_archivo,empresa)
                    os.remove(nombre_archivo)
                    time.sleep(5) # Esperamos 5 segundos para no saturar el servidor de correo
                else:
                    # Crear la carpeta "errores" si no existe
                    if not os.path.exists("errores"):
                        os.makedirs("errores")
                    # Mover el archivo de la nómina a la carpeta de errores con el nombre del DNI
                    nuevo_nombre = os.path.join("errores", f"{dni}.pdf")
                    os.rename(nombre_archivo, nuevo_nombre)   
            else:
                log(f"¡¡ERROR!! No se encuentra ningun DNI en nomina_{pagina + 1}.pdf")
                # Crear la carpeta "errores" si no existe
                if not os.path.exists("errores"):
                    os.makedirs("errores")
                # Mover el archivo de la nómina a la carpeta de errores con el nombre del DNI
                nuevo_nombre = os.path.join("errores", f"{dni}.pdf")
                os.rename(nombre_archivo, nuevo_nombre)   
else:
    log(f"¡¡ERROR!! El archivo {archivo_pdf} no existe")

