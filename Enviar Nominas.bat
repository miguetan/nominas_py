@echo off
echo Asegurate que tienes un archivo llamado nominas.pdf con las diferentes nóminas que deseas enviar un un archivo llamado emails.xls con los distintos emails asociados a los dni que aparecen en las nóminas ¿Desea continuar? (S/N)
set /p UserInput=
if /i "%UserInput%"=="S" (
    python3 .\script.py
) else (
    echo Operación abortada por el usuario.
)
pause