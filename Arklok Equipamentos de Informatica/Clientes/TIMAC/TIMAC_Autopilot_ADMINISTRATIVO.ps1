# DESABILITAR QUICK EDIT DO POWERSHELL (PARA MOUSE NÃO INTERROMPER EXECUÇÃO DO SCRIPT)
$CHECK1 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' | Select-Object -ExpandProperty QuickEdit
$CHECK2 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' | Select-Object -ExpandProperty QuickEdit

if (($CHECK1 -eq 1) -or ($CHECK2 -eq 1)) {
  Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000000

  Start-Process -FilePath .\Timac_Autopilot_ADMINISTRATIVO.exe
  exit
}

Start-Transcript -Path "C:\PerfLogs\Transcript.txt" -Force | Out-Null

$ErrorActionPreference = 'Stop'
Write-Host -ForegroundColor Green "Por gentileza, aguarde a instalação dos pacotes...`n"

$INSTALL_LOOP = $true;
$RETRIES_INSTALL = 0;

# INSTALAÇÃO DOS PACOTES E LOOP DE TENTATIVAS
while ($INSTALL_LOOP -eq $true) {
  try {
    Install-PackageProvider NuGet -Force -ErrorAction Stop;
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
    Install-Script -Name Get-WindowsAutoPilotInfo -Force -ErrorAction Stop
    $INSTALL_LOOP = $false;
  } catch {
    $RETRIES_INSTALL++
    if ($RETRIES_INSTALL -eq 2) {
      Start-Sleep -Seconds 2
      $key = ""

      while ($key -ne "r") {
        Write-Host -ForegroundColor Red "`nNão foi possível instalar os pacotes necessários, por gentileza verificar o status da internet."

        Write-Host -ForegroundColor Cyan -NoNewLine "`nPressione N para abrir as configurações da internet.";
        Write-Host -ForegroundColor Cyan -NoNewLine "`nPressione R para reiniciar o script.";
        Write-Host -ForegroundColor Cyan -NoNewLine "`nPressione E para encerrar a execução.`n";

        $key = [Console]::ReadKey($true).Key
        Clear-Host

        if ($key -eq "n") {
          Start-Process ms-settings:network
          ipconfig /renew
          Clear-Host
        } elseif ($key -eq "e") {
          exit
        }
      }

      Clear-Host
      $RETRIES_INSTALL = 0;
      Write-Host -ForegroundColor Green "Por gentileza, aguarde a instalação dos pacotes...`n"
      Continue;
    }
    Write-Host -ForegroundColor Yellow "`nProblema na instalação dos pacotes. Tentando novamente em alguns segundos...`n"
    Start-Sleep -Seconds 10
  }
}

# INSTRUÇÕES E FOMRATAÇÃO PARA USUÁRIOS
Write-Host -ForegroundColor Green "`nEm caso de erro, rodar o script novamente. Certifique-se de estar conectado à internet.`n"
Write-Host -ForegroundColor Green "Após a execução, o computador deve ser reiniciado.`n"

# LOOP DE LOGIN MICROSOFT
$MICROSOFT_LOGIN_LOOP = $true;

while ($MICROSOFT_LOGIN_LOOP -eq $true) {
  try {
    Get-WindowsAutoPilotInfo.ps1 -online -GroupTag "TAGS-TIMAC-AGRO-BRASIL" -Assign -ErrorAction Stop
    $MICROSOFT_LOGIN_LOOP = $false;
  } catch [System.Exception] {
    Write-Host -ForegroundColor Yellow "`nProblema ao logar na conta da Microsoft, tentando novamente em alguns segundos..."
    Start-Sleep -Seconds 10
  } Catch {
    Write-Host -ForegroundColor Yellow "`nProblema ao logar na conta da Microsoft, tentando novamente em alguns segundos..."
    Start-Sleep -Seconds 10
  }
}

Stop-Transcript | Out-Null

# VALIDAÇÃO DO PROCESSO DE IMPORT/ASSIGNMENT USANDO O TRANSCRIPT COMO BASE
$CHECK_IMPORT_SUCCESS = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "1 devices imported successfully"
$CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "ZtdDeviceAssignedToOtherTenant"
$CHECK_IF_ALREADY_ASSIGNED = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "ZtdDeviceAlreadyAssigned"
$CHECK_IF_ASSIGNED_0_SECONDS = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "Elapsed time to complete assignment: 0 seconds"

Write-Host -ForegroundColor Cyan "TAG utilizada: TAGS-TIMAC-AGRO-BRASIL`n"

if (!$CHECK_IMPORT_SUCCESS) { Write-Host -ForegroundColor Red "O processo de IMPORT não foi bem-sucedido."}
if ($CHECK_IF_ALREADY_ASSIGNED) { Write-Host -ForegroundColor Red "Máquina já possui perfil atrelado a TIMAC." }
if ($CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT) { Write-Host -ForegroundColor Red "Máquina já possui perfil atrelado em outro provedor." }
if ($CHECK_IF_ASSIGNED_0_SECONDS) { Write-Host -ForegroundColor Red "O processo de ASSIGNMENT não foi bem-sucedido." }

if (($CHECK_IMPORT_SUCCESS) -and !($CHECK_IF_ASSIGNED_0_SECONDS) -and !($CHECK_IF_ALREADY_ASSIGNED)) {
  Write-Host -ForegroundColor Green "Processo de IMPORT e ASSIGNMENT finalizados com sucesso."
}

Write-Host -ForegroundColor Cyan "`nValidar se o processo de IMPORT e ASSIGNMENT acima foram finalizados com sucesso antes de reiniciar a máquina.`n"

# Write-Host "`n"
# Write-Host -ForegroundColor Cyan "***************************************"
# Write-Host -ForegroundColor Cyan "* Criado por: Jan Seulin              *"
# Write-Host -ForegroundColor Cyan "* Setor: Imagem                       *"
# Write-Host -ForegroundColor Cyan "* Arklok Equipamentos de Informática  *"
# Write-Host -ForegroundColor Cyan "***************************************"
# Write-Host "`n"

$KEY_RESTART = ""

# REABILITAR QUICK EDIT DO POWERSHELL
Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000001

Write-Host "`n"

While ($KEY_RESTART -ne "r") {
  Write-Host -ForegroundColor Green "Pressione R para reiniciar a máquina agora."
  $KEY_RESTART = [Console]::ReadKey($true).Key

  if ($KEY_RESTART -eq "r") {
    Remove-Item -Path "C:\PerfLogs\Transcript.txt"
    Restart-Computer
  }
}