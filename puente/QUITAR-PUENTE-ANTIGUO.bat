@echo off
rem Quita el puente ANTIGUO del ordenador de administracion (Araceli), cuando el nuevo ya funcione.
schtasks /delete /tn "Copia RECLECHE" /f
if errorlevel 1 ( echo No estaba la tarea "Copia RECLECHE" o no se pudo quitar. ) else ( echo Quitada la copia antigua. )
pause
