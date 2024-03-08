# nominas_py
Este script recibe un pdf multihoja. lo divide en distintas hojas, extrae el dni asociado, busca en una base de datos el email asociado y envia la nomina con arhivo adjunto

Para su funcionamiento debe existir un archivo emails.xls. En la primera columna deben figurar los dnis y en la segunda columna los emails
La primera fila del archivo excel debe poner A1=dni B1=email

El archivo config.txt indica los datos de conexion con el servidor smtp.
La conexion debe ser por ssl (no tls)

El procesamiento se hace sobre un archivo llamado nominas.pdf que debe existir y contener tantas paginas como nominas se quieren enviar por correo