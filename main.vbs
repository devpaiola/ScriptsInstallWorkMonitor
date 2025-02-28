Option Explicit

' Objetos Globais
Dim objShell, fso, objNetwork, userName
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")
userName = objNetwork.UserName

Dim usuario
usuario = InputBox("Digite o nome de usuário:", "Credenciais de Login", userName)

If usuario = "" Then
    WScript.Quit
End If

' Verifica se o executável existe antes de rodar
Dim workExe
workExe = objShell.ExpandEnvironmentStrings("%APPDATA%\spa\work.exe")

If fso.FileExists(workExe) Then
    objShell.Run """" & workExe & """", 1, False
    WScript.Sleep 5000

    objShell.SendKeys "teste" ' Não alterar
    WScript.Sleep 500
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500
    objShell.SendKeys "teste"
    WScript.Sleep 500
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500
    objShell.SendKeys "tste"
    WScript.Sleep 500
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500
    objShell.SendKeys usuario
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}"
    WScript.Sleep 3000
Else
    MsgBox "Erro: O arquivo " & workExe & " não foi encontrado.", vbCritical, "Erro"
End If
