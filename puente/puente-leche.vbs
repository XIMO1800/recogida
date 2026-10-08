' PUENTE LECHE - pide a Google los ficheros nuevos o rehechos de la recogida y la recepcion de leche
' y los deja en el servidor:  R*.TXT -> RECLECHE   D*.TXT -> DESLECHE
' Lo lanza Windows cada 5 minutos (tarea "Puente leche"). Registro en envios.log, al lado de este fichero.
' La clave va en clave.txt (no se escribe aqui).
Option Explicit
Const URL = "https://script.google.com/macros/s/AKfycbx7uB5N55ZiEpH2EpCQBNsK-YI0X_7wVlkDvweUWNcBNplai0csOa34sMIEKptRdVRe/exec"
Dim BASES : BASES = Array("\\Servidor-i7\servidor\SERVIDORW10", "V:\SERVIDORW10")

Dim fso, dirScript, logPath, clave, base, equipo, b, lista, lineas, i, n, fallos
Set fso = CreateObject("Scripting.FileSystemObject")
dirScript = fso.GetParentFolderName(WScript.ScriptFullName)
logPath = dirScript & "\envios.log"
equipo = CreateObject("WScript.Network").ComputerName

clave = ""
If fso.FileExists(dirScript & "\clave.txt") Then clave = Trim(Replace(Replace(fso.OpenTextFile(dirScript & "\clave.txt", 1).ReadAll, vbCr, ""), vbLf, ""))
If clave = "" Then Apunta "ERROR: falta la clave en clave.txt" : WScript.Quit 1

base = ""
For Each b In BASES
  If fso.FolderExists(b & "\RECLECHE") Then base = b : Exit For
Next
If base = "" Then Apunta "ERROR: no se llega a la carpeta RECLECHE del servidor" : WScript.Quit 1

On Error Resume Next
lista = PedirTexto("tipo=pendientes")
If Err.Number <> 0 Then Apunta "ERROR de conexion con Google: " & Err.Description : WScript.Quit 1
On Error GoTo 0
lineas = Split(Replace(lista, vbCr, ""), vbLf)
If UBound(lineas) < 0 Then Apunta "ERROR: respuesta vacia de Google" : WScript.Quit 1
If lineas(0) <> "OK" Then Apunta "ERROR de Google: " & Left(lista, 200) : WScript.Quit 1

n = 0 : fallos = 0
For i = 1 To UBound(lineas)
  If Trim(lineas(i)) <> "" Then
    If Entregar(lineas(i)) Then n = n + 1 Else fallos = fallos + 1
  End If
Next
WScript.Quit fallos

' ---------------------------------------------------------------------------------------------

Function Entregar(linea)
  Dim p, nombre, destino, version, carpeta, h, cuerpo, tmp, final, st, habia, re
  Entregar = False
  p = Split(linea, ";")
  If UBound(p) < 2 Then Apunta "ERROR: linea rara: " & linea : Exit Function
  nombre = UCase(Trim(p(0))) : destino = UCase(Trim(p(1))) : version = Trim(p(2))
  Set re = New RegExp : re.Pattern = "^[RD][0-9]{6}[0-9A-Z]\.TXT$"
  If Not re.Test(nombre) Then Apunta "ERROR: nombre no valido: " & nombre : Exit Function
  If destino <> "RECLECHE" And destino <> "DESLECHE" Then Apunta "ERROR: carpeta no valida: " & destino : Exit Function
  carpeta = base & "\" & destino
  On Error Resume Next
  If Not fso.FolderExists(carpeta) Then fso.CreateFolder carpeta
  Set h = Pedir("tipo=fichero&nombre=" & nombre)
  If Err.Number <> 0 Then Apunta "ERROR al bajar " & nombre & ": " & Err.Description : Exit Function
  If Left(h.responseText, 6) = "ERROR;" Then Apunta "ERROR de Google con " & nombre & ": " & h.responseText : Exit Function
  cuerpo = h.responseBody
  If LenB(cuerpo) = 0 Then Apunta "ERROR: " & nombre & " ha llegado vacio" : Exit Function
  ' Primero con nombre temporal y luego se renombra: el programa nunca ve un fichero a medio escribir
  tmp = carpeta & "\" & Left(nombre, Len(nombre) - 4) & ".TMP"
  final = carpeta & "\" & nombre
  Set st = CreateObject("ADODB.Stream")
  st.Type = 1 : st.Open : st.Write cuerpo : st.SaveToFile tmp, 2 : st.Close
  If Err.Number <> 0 Then Apunta "ERROR al escribir " & tmp & ": " & Err.Description : Exit Function
  habia = fso.FileExists(final)
  If habia Then fso.DeleteFile final, True
  fso.MoveFile tmp, final
  If Err.Number <> 0 Then Apunta "ERROR al dejar " & final & ": " & Err.Description : Exit Function
  ' Solo despues de dejarlo se confirma a Google
  Dim ok : ok = PedirTexto("tipo=entregado&nombre=" & nombre & "&version=" & version & "&equipo=" & equipo)
  If Err.Number <> 0 Or ok <> "OK" Then Apunta "AVISO: " & nombre & " dejado en " & destino & " pero no se pudo confirmar (se repetira): " & Err.Description & ok : Exit Function
  On Error GoTo 0
  If habia Then
    Apunta "rehecho " & nombre & " en " & destino & " (machaca el anterior)"
  Else
    Apunta "copiado " & nombre & " a " & destino
  End If
  Entregar = True
End Function

Function Pedir(q)
  Dim h : Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  h.setTimeouts 15000, 15000, 30000, 60000
  h.open "GET", URL & "?" & q & "&clave=" & clave & "&t=" & Int(Timer * 100), False
  h.send
  If h.status <> 200 Then Err.Raise vbObjectError + 1, "Puente", "HTTP " & h.status
  Set Pedir = h
End Function

Function PedirTexto(q)
  PedirTexto = Pedir(q).responseText
End Function

Sub Apunta(t)
  On Error Resume Next
  If fso.FileExists(logPath) Then
    If fso.GetFile(logPath).Size > 1000000 Then
      If fso.FileExists(dirScript & "\envios.old.log") Then fso.DeleteFile dirScript & "\envios.old.log", True
      fso.MoveFile logPath, dirScript & "\envios.old.log"
    End If
  End If
  Dim f : Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Date & " " & Time & "  " & t
  f.Close
End Sub
