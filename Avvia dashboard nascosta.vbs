Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
sourceDir = fso.GetParentFolderName(WScript.ScriptFullName)
localAppData = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
vendorRoot = localAppData & "\NordLaser"
localRoot = vendorRoot & "\DichiarazioniOrigine"
appDir = localRoot & "\app"
logFile = localRoot & "\dashboard.log"
portFile = localRoot & "\dashboard.port"

If Not fso.FolderExists(vendorRoot) Then
  fso.CreateFolder(vendorRoot)
End If

If Not fso.FolderExists(localRoot) Then
  fso.CreateFolder(localRoot)
End If

syncCmd = "cmd.exe /d /c robocopy """ & sourceDir & """ """ & appDir & _
  """ /MIR /FFT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP /XD __pycache__ >nul"
syncResult = shell.Run(syncCmd, 0, True)

If syncResult > 7 Then
  MsgBox "Impossibile copiare la dashboard sul computer. Verificare l'accesso alla cartella di rete.", 16, "Dichiarazioni origine"
  WScript.Quit
End If

portablePy = appDir & "\runtime\python.exe"
If Not fso.FileExists(portablePy) Then
  MsgBox "Runtime portatile mancante. Contattare il responsabile della dashboard.", 16, "Dichiarazioni origine"
  WScript.Quit
End If

If fso.FileExists(portFile) Then
  On Error Resume Next
  Set portStream = fso.OpenTextFile(portFile, 1)
  existingPort = Trim(portStream.ReadAll)
  portStream.Close
  Set request = CreateObject("MSXML2.XMLHTTP")
  request.Open "GET", "http://127.0.0.1:" & existingPort & "/", False
  request.Send
  If Err.Number = 0 And request.Status = 200 Then
    shell.Run "http://127.0.0.1:" & existingPort & "/", 1, False
    WScript.Quit
  End If
  Err.Clear
  On Error GoTo 0
  fso.DeleteFile portFile, True
End If

cmd = "cmd.exe /d /c cd /d """ & appDir & """ && runtime\python.exe server.py >> """ & logFile & """ 2>&1"
shell.Run cmd, 0, False

For attempt = 1 To 60
  WScript.Sleep 250
  If fso.FileExists(portFile) Then
    Set portStream = fso.OpenTextFile(portFile, 1)
    port = Trim(portStream.ReadAll)
    portStream.Close
    If Len(port) > 0 Then
      shell.Run "http://127.0.0.1:" & port & "/", 1, False
      WScript.Quit
    End If
  End If
Next

MsgBox "La dashboard non si e avviata. Consultare il file:" & vbCrLf & logFile, 16, "Dichiarazioni origine"
