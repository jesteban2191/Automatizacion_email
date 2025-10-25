@echo off

:: Crear el entorno virtual
echo Creando entorno virtual...
call python -m venv .venv
echo Entorno virtual creado satisfactoriamente

echo.
:: Activar el entorno virtual
echo Activando entorno virtual...
call .venv\Scripts\activate
echo.
echo Entorno virtual activado

echo.
:: Actualizando el PIP
echo Actualizando el PIP...
call python -m pip install --upgrade pip
echo.
echo PIP actualizado satisfactoriamente

echo.
:: Instalar las bibliotecas necesarias
echo Instalando bibliotecas...
call pip install -r requirements.txt
echo.
echo Librerias instaladas satisfactoriamente

echo.
:: Mostrar un mensaje de finalización
echo Entorno virtual creado, activado y librerias instaladas correctamente.
echo.
call deactivate
pause

