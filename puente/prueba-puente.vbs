' PRUEBA DEL PUENTE LECHE - comprueba en ESTE ordenador la conexion con Google y con el servidor. No copia nada.
Option Explicit
Const URL = "https://script.google.com/macros/s/AKfycbx7uB5N55ZiEpH2EpCQBNsK-YI0X_7wVlkDvweUWNcBNplai0csOa34sMIEKptRdVRe/exec"
Dim fso, dirScript, clave, h, r1, r2, r3, b, t, base
Set fso = CreateObject("Scripting.FileSystemObject")
dirScript = fso.GetParentFolderName(WScript.ScriptFullName)
clave = "x"
If fso.FileExists(dirScript & "\clave.txt") Then clave = Trim(Replace(Replace(fso.OpenTextFile(dirScript & "\clave.txt", 1).ReadAll, vbCr, ""), vbLf, ""))

On Error Resume Next
Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
h.setTimeouts 15000, 15000, 30000, 60000
h.open "GET", URL & "?tipo=pendientes&clave=" & clave, False
h.send
If Err.Number <> 0 Then
  r1 = "MAL: " & Err.Description & vbCrLf & "(si habla de seguridad o certificados, a este Windows 7 le faltan las actualizaciones de TLS 1.2)"
ElseIf h.status <> 200 Then
  r1 = "MAL: HTTP " & h.status
Else
  t = h.responseText
  If Left(t, 2) = "OK" Then
    r1 = "BIEN. Clave correcta. Ficheros pendientes: " & (UBound(Split(Replace(t, vbCr, ""), vbLf)))
  ElseIf InStr(t, "clave") > 0 Then
    r1 = "BIEN (la conexion funciona). La clave todavia no esta puesta o no es correcta."
  Else
    r1 = "BIEN (la conexion funciona). Respuesta: " & Left(t, 120)
  End If
End If
Err.Clear

base = ""
For Each b In Array("\\Servidor-i7\servidor\SERVIDORW10", "V:\SERVIDORW10")
  If fso.FolderExists(b & "\RECLECHE") Then base = b : Exit For
Next
If base = "" Then
  r2 = "MAL: no se ve \\Servidor-i7\servidor\SERVIDORW10\RECLECHE"
Else
  fso.CreateTextFile(base & "\RECLECHE\PRUEBAPT.TMP", True).Close
  If Err.Number <> 0 Then
    r2 = "MAL: se ve " & base & " pero no deja escribir: " & Err.Description
  Else
    fso.DeleteFile base & "\RECLECHE\PRUEBAPT.TMP", True
    r2 = "BIEN: " & base & " (se puede escribir)"
  End If
End If
MsgBox "Conexion con Google:" & vbCrLf & r1 & vbCrLf & vbCrLf & "Servidor:" & vbCrLf & r2, 64, "Prueba del puente leche"
