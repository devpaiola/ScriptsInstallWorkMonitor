Option Explicit

'Nossos OBJS Globais
Dim objShell, fso, objNetwork, userName
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")
userName = objNetwork.UserName 

Dim opcao 'Case
Do
    opcao = InputBox( _
        "Gerenciador WorkMonitor" & vbNewLine & vbNewLine & _
        "Digite a opção desejada:" & vbNewLine & _
        "1 - Instalar" & vbNewLine & _
        "2 - Desinstalar" & vbNewLine & _
        "3 - Sair", _
        "Controle WorkMonitor", _
        "1" _
    )
    
    If opcao = "" Then  ' Usuário clicou em Cancelar
        WScript.Quit
    ElseIf IsNumeric(opcao) Then
        opcao = CInt(opcao)
        Select Case opcao
            Case 1: Call InstalarWorkMonitor : Exit Do
            Case 2: Call DesinstalarWorkMonitor : Exit Do
            Case 3: WScript.Quit
            Case Else: MsgBox "Opcao invalida. Tente novamente.", vbExclamation, "Erro"
        End Select
    Else
        MsgBox "Entrada invalida. Use números de 1 a 3.", vbExclamation, "Erro"
    End If
Loop


Sub InstalarWorkMonitor()
    Dim objShell, fso, scriptPath, file, regFile, flagFile, installPath, folder
    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
    flagFile = objShell.ExpandEnvironmentStrings("%temp%\install_done.flag")
    installPath = objShell.ExpandEnvironmentStrings("%temp%\install.exe")
    
    ' Verifica se a instalação já foi realizada
    If fso.FileExists(flagFile) Then
        WScript.Echo "Instalacoo ja realizada."
        Exit Sub
    End If
    
    ' Executa todos os arquivos .reg na mesma pasta do script
    Set folder = fso.GetFolder(scriptPath)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "reg" Then
            objShell.Run "regedit /s """ & file.Path & """", 0, True
        End If
    Next
    
    ' Baixa o instalador (se necessário)
    Call BaixarInstalador
    If Not fso.FileExists(installPath) Then
        WScript.Echo "Erro: O instalador não foi baixado corretamente."
        WScript.Quit
    End If
    
    ' Executa o instalador
    objShell.Run """" & installPath & """ /s /v ""/qn""", 0, True
    
    ' Cria um arquivo de flag indicando que a instalação foi concluída
    Dim flag
    Set flag = fso.CreateTextFile(flagFile, True)
    flag.WriteLine "Instalacao concluida."
    flag.Close
    
    ' Se nenhum arquivo .reg foi encontrado, chama PreencherCredenciais
    Dim regFound: regFound = False
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "reg" Then
            regFound = True
            Exit For
        End If
    Next
    
    If Not regFound Then
        Call PreencherCredenciais
    End If
    
    ' Continua o processo de configuração
    Call IngressarNoCanal
    Call ConfigurarEnergia
    
    WScript.Echo "Instalacao e configuracao concluidas."
End Sub



Sub DesinstalarWorkMonitor()
    Dim uninstallBat, batContent
    uninstallBat = objShell.ExpandEnvironmentStrings("%TEMP%\uninstall.bat")
    
    batContent = "@echo off" & vbCrLf & _
    "echo Desinstalando e removendo arquivos e registros..." & vbCrLf & _
    "taskkill /F /IM spa.exe /IM xspa.exe /IM work.exe " & vbCrLf & _
    "if exist ""%appdata%\spa"" rmdir /S /Q ""%appdata%\spa""" & vbCrLf & _
    "if exist ""%temp%\install_done.flag"" del /Q ""%temp%\install_done.flag""" & vbCrLf & _
    "if exist ""%temp%\install.exe"" del /Q ""%temp%\install.exe""" & vbCrLf & _
    "reg delete ""HKEY_CURRENT_USER\Software\SmartComputadores"" /f" & vbCrLf & _
    "echo Processo concluido." & vbCrLf & _
    "pause"
    
    With fso.CreateTextFile(uninstallBat, True)
        .Write batContent
        .Close
    End With
    
    objShell.Run "cmd /c """ & uninstallBat & """", 0, True    
    WScript.Sleep 5000
    If fso.FileExists(uninstallBat) Then fso.DeleteFile(uninstallBat)
End Sub









'Classes Personalizaveis
Sub BaixarInstalador ()
    Dim installPath
    installPath = objShell.ExpandEnvironmentStrings("%temp%\install.exe")
    objShell.Run "cmd /c curl -o """ & installPath & """ https://www.workmonitor.com/install/install.exe", 0, True
End Sub

Sub ImportarRegistros()
     Dim regFolder, file, foundRegFiles
    Dim regFiles
    regFiles = objShell.CurrentDirectory & "\*.reg"

    foundRegFiles = False
    Set regFolder = fso.GetFolder(objShell.CurrentDirectory)
    For Each file In regFolder.Files
        If LCase(fso.GetExtensionName(file)) = "reg" Then
            foundRegFiles = True
            WScript.Echo "Importando: " & file.Name
            objShell.Run "reg import """ & file.Path & """", 0, True
            If objShell.Exec("reg import """ & file.Path & """").ExitCode <> 0 Then
                WScript.Echo "Erro ao importar " & file.Name
            Else
                WScript.Echo file.Name & " importado com sucesso."
            End If
        End If
    Next
    If Not foundRegFiles Then
        WScript.Echo "Nenhum arquivo .reg encontrado."
    End If
End Sub

Sub PreencherCredenciais()
    Dim usuario, dominio, senha
    usuario = InputBox("Digite o nome de usuario:", "Credenciais de Login", userName)
    If usuario = "" Then Exit Sub

    objShell.Run """%APPDATA%\spa\work.exe""", 1, False
    WScript.Sleep 5000 
    objShell.SendKeys "teste" ' Não alterar
    WScript.Sleep 500
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500
    objShell.SendKeys "teste"
    WScript.Sleep 500
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500
    objShell.SendKeys "teste123"
    WScript.Sleep 500
    objShell.SendKeys "{TAB}"
    WScript.Sleep 500
    objShell.SendKeys usuario
    WScript.Sleep 500
    objShell.SendKeys "{ENTER}" 
    WScript.Sleep 3000
End Sub

Sub IngressarNoCanal()
    WScript.Sleep 5000
    objShell.Run """%APPDATA%\spa\work.exe"" /ingressar", 0, True
End Sub

Sub ConfigurarEnergia()
    Dim psCommand
    psCommand = "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -NoProfile -Command " & _
                """powercfg /s SCHEME_BALANCED; " & _
                "powercfg /change disk-timeout-ac 0; " & _
                "powercfg /change disk-timeout-dc 0; " & _
                "powercfg -SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0; " & _
                "powercfg -SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0; " & _
                "powercfg -SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 96996bc0-ad50-47ec-923b-6f41874dd9eb 0; " & _
                "powercfg -SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 96996bc0-ad50-47ec-923b-6f41874dd9eb 0; " & _
                "powercfg -SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 3; " & _
                "powercfg -SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 3; " & _
                "powercfg -SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e SUB_VIDEO VIDEOIDLE 0; " & _
                "powercfg -SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e SUB_VIDEO VIDEOIDLE 0; " & _
                "powercfg -SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e SUB_SLEEP STANDBYIDLE 0; " & _
                "powercfg -SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e SUB_SLEEP STANDBYIDLE 0; " & _
                "powercfg /apply""" & _
                "' -Verb RunAs"
                
    objShell.Run "powershell -Command """ & psCommand & """", 0, True
    Dim spaPath
    spaPath = objShell.ExpandEnvironmentStrings("%APPDATA%\spa")
    objShell.Run "powershell -Command ""Start-Process powershell -ArgumentList 'Add-MpPreference -ExclusionPath """ & spaPath & """' -Verb RunAs""", 0, True
End Sub