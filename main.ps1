$global:tempPath = "$env:TEMP"
$global:installFlag = "$tempPath\install_done.flag"
$global:installExe = "$tempPath\install.exe"
$global:vbsScript = "$tempPath\install.vbs"

function Show-Menu {
    do {
        $option = Read-Host "Gerenciador WorkMonitor

Digite a opcao desejada:
1 - Instalar
2 - Desinstalar
3 - Sair"
        
        switch ($option) {
            "1" { Install-WorkMonitor; break }
            "2" { Uninstall-WorkMonitor; break }
            "3" { exit }
            default { Write-Host "Opcao invalida. Tente novamente." -ForegroundColor Red }
        }
    } while ($true)
}

function Install-WorkMonitor {
    if (Test-Path $installFlag) {
        Write-Host "Instalacao ja realizada."
        return
    }
    
    DownloadInstaller

    if (Test-Path $installExe) {
        Install-Registry
        
        
        Start-Process -FilePath $installExe -ArgumentList @('/s', '/v', '/qn') -NoNewWindow 
        Install-VBS
        New-Item -Path $installFlag -ItemType File -Force | Out-Null
        Write-Host "Instalacao e configuracao concluidas."
        exit
    } else {
        Write-Host "Erro: O instalador nao foi baixado corretamente." -ForegroundColor Red
    }
}

function Uninstall-WorkMonitor {
    Write-Host "Desinstalando e removendo arquivos e registros..."
    Stop-Process -Name "spa" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "xspa" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "work" -Force -ErrorAction SilentlyContinue

    Remove-Item -Path "$env:APPDATA\spa" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $installFlag -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $installExe -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "HKCU:\Software\SmartComputadores" -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "Processo concluido."
}

function DownloadInstaller {
    Write-Host "Baixando instalador..."
    try {
        Invoke-WebRequest -Uri "https://www.workmonitor.com/install/install.exe" -OutFile $installExe

    } catch {
        Write-Host "Erro ao baixar o instalador: $_" -ForegroundColor Red
    }
}

function Install-Registry {
    $regFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.reg"
    if ($regFiles.Count -gt 0) {
        foreach ($file in $regFiles) {
            reg import $file.FullName
            if ($?) {
                Write-Host "$($file.Name) importado com sucesso."
            } else {
                Write-Host "Erro ao importar $($file.Name)." -ForegroundColor Red
            }
        }
        return $true
    } else {
        Write-Host "Nenhum arquivo .reg encontrado."
        return $false
        exit\b
    }
}

function Install-VBS {
    Write-Host "Baixando script VBS..."
    try {
        $vbsUrl = "https://raw.githubusercontent.com/devpaiola/ScriptsInstallWorkMonitor/refs/heads/main/InstaladorFab/FinalTEste.vbs"
        Invoke-WebRequest -Uri $vbsUrl -OutFile $vbsScript
        if (Test-Path $vbsScript) {
            Write-Host "Executando script VBS..."
            Start-Process "wscript" -ArgumentList "`"$vbsScript`"" -NoNewWindow -Wait
        } else {
            Write-Host "Erro ao baixar o script VBS." -ForegroundColor Red
        }
    } catch {
        Write-Host "Erro ao baixar o script: $_" -ForegroundColor Red
    }
}

function ConfigurePowerSettings {           #Adicionar o script de configs de energia 
    powercfg /s SCHEME_BALANCED
    powercfg /change disk-timeout-ac 0
    powercfg /change disk-timeout-dc 0
    powercfg -SETACVALUEINDEX SCHEME_BALANCED SUB_SLEEP STANDBYIDLE 0
    powercfg -SETDCVALUEINDEX SCHEME_BALANCED SUB_SLEEP STANDBYIDLE 0
    powercfg /apply
    
    $spaPath = "$env:APPDATA\spa"
    Add-MpPreference -ExclusionPath $spaPath
}

Show-Menu
