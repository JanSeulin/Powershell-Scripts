$host.ui.RawUI.WindowTitle = "MDIAS AUTOPILOT"

# FUNÇÃO PARA MOSTRAR CRÉDITOS NO FIM DA EXECUÇÃO
function RunCredits {
  Write-Host -ForegroundColor Cyan "**************************************"
  Start-Sleep -Milliseconds 250
  Write-Host -ForegroundColor Cyan "*                                    *"
  Start-Sleep -Milliseconds 250
  Write-Host -ForegroundColor Cyan "* Criado por: Jan Seulin             *"
  Start-Sleep -Milliseconds 250
  Write-Host -ForegroundColor Cyan "* Setor: Imagem                      *"
  Start-Sleep -Milliseconds 250
  Write-Host -ForegroundColor Cyan "* Arklok Equipamentos de Informática *"
  Start-Sleep -Milliseconds 250
  Write-Host -ForegroundColor Cyan "*                                    *"
  Start-Sleep -Milliseconds 250
  Write-Host -ForegroundColor Cyan "**************************************"
  Start-Sleep -Milliseconds 250
  Write-Host "`n"
}


# DESABILITAR QUICK EDIT DO POWERSHELL
$CHECK1 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' | Select-Object -ExpandProperty QuickEdit
$CHECK2 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' | Select-Object -ExpandProperty QuickEdit

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

$ARRAY_TAGS=@('MDBN','MDBD','GMEN','GMED','MOIN','MOID','MRCN','MRCD','CRTN','CRTD','GMPN','GMPD','SJMN','SJMD','TMNN','TMND','DMAN','DMAD','DPIN','DPID','BLMN','BLMD','EMAN','EMAD','GMTN','GMTD','JPAN','JPAD','VITN','VITD','VITN','VITD','MAC','MACD','AJUN','AJUD','VDCN','VDCD','GMAN','GMAD','GMAN','GMAD','RJN','RJD','PEN','PED','PMN','PMD','PQN','PQD','UBLN','UBLD','BSAN','BSAD','SCSN','SCSD','GRUN','GRUD','LPAN','LPAD','MPRN','MPRD','LTXN','LTXD','JASN','JASD','BGO','BGOD','PNSN','PNSD','BGON','BGOD', 'MADN', 'MADD', 'MADNW', 'MADDW')

$TAG_CORRECT = $false;

while (!$TAG_CORRECT) {
  $TAG = Read-Host "Digite a TAG da unidade"

  while ($TAG.length -lt 3){
    $TAG = Read-Host 'A TAG deve ter no mínimo 3 caracteres.'
  }

  $TAG = ($TAG -replace '\s', '').ToUpper()
  Write-Host -ForegroundColor Blue "`nTAG SELECIONADA: $TAG`n"

  if (!$ARRAY_TAGS.Contains($TAG)) {
    Write-Host -Foregroundcolor Yellow "TAG não localizada. Por gentileza, validar se a TAG está correta.`n"
  } else {
    $TAG_CORRECT = $true;
    Write-Host -ForegroundColor Green "TAG válida, aguarde...`n"
  }
}

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
if ($CHECK_IF_ALREADY_ASSIGNED) { Write-Host -ForegroundColor Red "Máquina já possui perfil na MDIAS." }
if ($CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT) { Write-Host -ForegroundColor Red "Máquina já possui perfil de Autopilot em outro provedor." }

if (($CHECK_IMPORT_SUCCESS) -and !($CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT) -and !($CHECK_IF_ALREADY_ASSIGNED)) {
  Write-Host -ForegroundColor Green "`nProcesso de importação finalizado com sucesso."
}

Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000001

Write-Host "`n"

if (($CHECK_IMPORT_SUCCESS) -and !($CHECK_IF_ALREADY_ASSIGNED_OTHER_TENANT) -and !($CHECK_IF_ALREADY_ASSIGNED)) {
  Write-Host -ForegroundColor Green "Processo de importação finalizado com sucesso.`n"

  net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

  & "\\172.18.3.4\d$\Servidor Deployment\MDT01\Scripts\Gerar_Log\Gerar_Log_Autopilot.exe" 'MDIAS (AUTOPILOT)' $TAG

  $KEY_RESTART = ""

  Write-Host "`n"
  While ($KEY_RESTART -ne "e") {

    Write-Host -ForegroundColor Green "Pressione E para encerrar o script."
    $KEY_RESTART = [Console]::ReadKey($true).Key

    if ($KEY_RESTART -eq "e") {
      Clear-Host
      RunCredits

      Write-Host "Encerrando..."
      Start-Sleep -Seconds 5

      Remove-Item -Path "C:\PerfLogs\Transcript.txt"
    }
  }
} else {
  Write-Host -ForegroundColor Red "Não foi possível realizar o procedimento necessário, por gentileza analisar os erros acima."

  $KEY_FAIL = ""

  Write-Host "`n"
  while ($KEY_FAIL -ne "r") {
    Write-Host -ForegroundColor Yellow "Pressione R para reiniciar o script."
    Write-Host "`n"
    $KEY_FAIL = [Console]::ReadKey($true).Key

    if ($KEY_FAIL -eq "r") {
      Clear-Host
      Remove-Item -Path "C:\PerfLogs\Transcript.txt"
      RunCredits

      Write-Host "Reiniciando Script..."
      Start-Sleep -Seconds 5
      Start-Process -FilePath .\Script_Autopilot.exe
      exit
    }
  }
}