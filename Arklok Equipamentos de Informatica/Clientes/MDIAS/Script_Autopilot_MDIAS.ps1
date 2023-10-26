# DESABILITAR QUICK EDIT DO POWERSHELL
$CHECK1 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' | Select-Object -ExpandProperty QuickEdit
$CHECK2 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' | Select-Object -ExpandProperty QuickEdit

Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Provisioning\Diagnostics\Autopilot | Select-Object -ExpandProperty CloudAssignedTenantDomain

if (($CHECK1 -eq 1) -or ($CHECK2 -eq 1)) {
  Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000000

  Start-Process -FilePath .\Script_Autopilot.exe
  exit
}

Start-Transcript -Path "C:\PerfLogs\Transcript.txt" -Force | Out-Null

$ErrorActionPreference = 'Stop' #Forçar erros a serem tratados como "erros terminais" para que seja possível a captura no bloco try-catch
Write-Host -ForegroundColor Green "Por gentileza, aguarde a instalação dos pacotes..."
# Write-Host -ForegroundColor Green "Pressione ENTER se solicitado qualquer confirmação durante esse processo.`n"

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
          Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000001
          Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000001
          Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000001
          exit
        }
      }

      Clear-Host
      $RETRIES_INSTALL = 0;
      Write-Host -ForegroundColor Green "Por gentileza, aguarde a instalação dos pacotes...`n"
      # Write-Host -ForegroundColor Green "Pressione ENTER se solicitado qualquer confirmação durante esse processo.`n"
      Continue;
    }
    Write-Host -ForegroundColor Yellow "`nProblema na instalação dos pacotes. Tentando novamente em alguns segundos...`n"
    Start-Sleep -Seconds 10
  }
}

Write-Output "`n"
Write-Host -ForegroundColor Green "Após inserir a TAG, aguarde a abertura da tela de login da Microsoft."
Write-Host -ForegroundColor Green "Em caso de erro, por gentileza rodar o script novamente. Certifique-se de estar conectado à internet.`n"

$TAG = Read-Host 'Digite a TAG da unidade'

while ($TAG.length -lt 3){
	$TAG = Read-Host 'A TAG deve ter no mínimo 3 caracteres. Digite a TAG da unidade'
}

$TAG = ($TAG -replace '\s', '').ToUpper()
Write-Host -ForegroundColor Blue "TAG SELECIONADA: $TAG"

$MICROSOFT_LOGIN_LOOP = $true;

while ($MICROSOFT_LOGIN_LOOP -eq $true) {
  try {
    Get-WindowsAutoPilotInfo -Online -GroupTag $TAG -ErrorAction Stop
    $MICROSOFT_LOGIN_LOOP = $false;

  } catch [System.Exception] {
    Write-Host -ForegroundColor Yellow "`nProblema ao logar na conta da Microsoft, tentando novamente em alguns segundos..."
    Start-Sleep -Seconds 10
  } Catch {
    Write-Host -ForegroundColor Yellow "`nProblema ao logar na conta da Microsoft, tentando novamente em alguns segundos..."
    Start-Sleep -Seconds 10
  }
}

# VALIDAÇÃO
Stop-Transcript | Out-Null
$CHECK_IMPORT_SUCCESS = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "1 devices imported successfully"
$CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "ZtdDeviceAssignedToOtherTenant"
$CHECK_IF_ALREADY_ASSIGNED = Select-String -Path "C:\PerfLogs\Transcript.txt" -Pattern "ZtdDeviceAlreadyAssigned"

Write-Host "`n"

if (!$CHECK_IMPORT_SUCCESS) { Write-Host -ForegroundColor Red "O processo de Importação não foi bem-sucedido."}
if ($CHECK_IF_ALREADY_ASSIGNED) { Write-Host -ForegroundColor Red "Máquina já possui perfil importado na MDIAS." }
if ($CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT) { Write-Host -ForegroundColor Red "Máquina já possui perfil atrelado em outro provedor." }

if (($CHECK_IMPORT_SUCCESS) -and !($CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT) -and !($CHECK_IF_ALREADY_ASSIGNED)) {
  Write-Host -ForegroundColor Green "`nProcesso de importação finalizado com sucesso."
}

Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000001

$KEY_CLOSE = ""
Write-Host "`n"

While ($KEY_CLOSE -ne "C") {
  Write-Host -ForegroundColor Green "Pressione C para encerrar a execução."
  $KEY_CLOSE = [Console]::ReadKey($true).Key

  if ($KEY_CLOSE -eq "C") {
    exit
  }
}