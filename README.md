# nominas_py
Este script recibe un pdf multihoja. lo divide en distintas hojas, extrae el dni asociado, busca en una base de datos el email asociado y envia la nomina con arhivo adjunto

Para su funcionamiento debe existir un archivo emails.xlsx. En la primera columna deben figurar los dnis y en la segunda columna los emails
La primera fila del archivo excel debe poner A1=dni B1=email
El programa buscará los emails en dos hojas distintas del libro. El nombre de la hoja indica el nombre de la empresa. Importante ya que ese nombre
indica desde que remitente de correo se enviará la nomina buscando sus datos de configuración el config.txt

El archivo config.txt indica los datos de conexion con el servidor smtp.


La conexion debe ser por ssl (no tls)

El procesamiento se hace sobre un archivo llamado nominas.pdf que debe existir y contener tantas paginas como nominas se quieren enviar por correo

Para lanzarlo debemos ejecutar el archivo por lotes 'Enviar Nomianas.bat'


REQUISITOS
Python 3
Instalar pyPDF2 y pandas
python3 -m pip install pandas
python3 -m pip install openpyxl
pip3 install pyPDF2
