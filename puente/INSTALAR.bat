@echo off
rem Instala el PUENTE LECHE en este ordenador (el del despacho, siempre encendido).
set "DIR=C:\PUENTE_LECHE"
if not exist "%DIR%" md "%DIR%"
rem "type" copia el texto sin la marca de "descargado de Internet", asi Windows no pregunta al abrirlo
type "%~dp0puente-leche.vbs" > "%DIR%\puente-leche.vbs"
type "%~dp0prueba-puente.vbs" > "%DIR%\prueba-puente.vbs"
if exist "%DIR%\clave.txt" goto tarea
echo.
set "CLAVE="
set /p CLAVE=Pega la CLAVE DEL PUENTE (la que sale al ejecutar instalarPuente) y pulsa Intro: 
if "%CLAVE%"=="" ( echo No has puesto la clave. & pause & exit /b )
>"%DIR%\clave.txt" echo %CLAVE%
:tarea
schtasks /create /tn "Puente leche" /tr "wscript.exe //B \"%DIR%\puente-leche.vbs\"" /sc minute /mo 5 /f
if errorlevel 1 ( echo ERROR: no se pudo crear la tarea programada. & pause & exit /b )
echo.
echo Primera pasada...
wscript.exe //B "%DIR%\puente-leche.vbs"
if exist "%DIR%\envios.log" type "%DIR%\envios.log"
echo.
echo Listo. Cada 5 minutos se dejaran en el servidor los ficheros nuevos de recogida (RECLECHE) y recepcion (DESLECHE).
echo Registro: %DIR%\envios.log
pause
