$host.ui.RawUI.WindowTitle = "BAT - EXTRAIR HASH"

# DESABILITAR QUICK EDIT DO POWERSHELL
$CHECK1 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' | Select-Object -ExpandProperty QuickEdit
$CHECK2 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' | Select-Object -ExpandProperty QuickEdit

if (($CHECK1 -eq 1) -or ($CHECK2 -eq 1)) {
  Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000000

  Start-Process -FilePath .\BAT_Extrair_HASH.exe
  exit
}

$PATH = Test-Path -Path "C:\HWID"
if (-Not $PATH){
	New-Item -Type Directory -Path "C:\HWID"
}

Set-Location -Path "C:\HWID"

$INSTALL_LOOP = $true;
$RETRIES_INSTALL = 0;

Write-Host -ForegroundColor Green "`nPor gentileza, aguarde a instalação dos pacotes..."
Write-Host -ForegroundColor Green "Pressione ENTER se solicitado qualquer confirmação durante esse processo.`n"

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

      Continue;
    }
    Write-Host -ForegroundColor Yellow "`nProblema na instalação dos pacotes. Tentando novamente em alguns segundos...`n"
    Start-Sleep -Seconds 10
  }
}

Write-Host "`n"

$PATRIMONIO = ""

while ($PATRIMONIO.length -lt 4 ) {
  $PATRIMONIO = Read-Host 'Digite o patrimônio da máquina'
  $PATRIMONIO = $PATRIMONIO -replace '\s', ''

  if ($PATRIMONIO.length -lt 4) { Write-Host -ForegroundColor Yellow "O patrimônio deve conter no mínimo 4 dígitos."}
}

$STRINGHWID = "AutoPilotHWID" + $PATRIMONIO + ".csv"
Get-WindowsAutoPilotInfo.ps1 -OutputFile $STRINGHWID

$DATE = Get-Date -Format "dd-MM"
# $TEST_HASH_FOLDER = Test-Path "$PENDRIVE_DRIVE_LETTER\HASH $DATE"

# NETWORK SHARE
$net = new-object -ComObject WScript.Network
$net.MapNetworkDrive("u:", "\\172.18.3.4\d$", $false, "arkserv\jan", "Lucy@505")

$HASH_PATH = "U:\SERVIDOR DEPLOYMENT\MDT01\HASH"
$HASH_TEST_PATH = Test-Path $HASH_PATH
$HASH_PATH_FOLDER = "$HASH_PATH\$DATE"
$HASH_PATH_FOLDER_TEST = Test-Path $HASH_PATH_FOLDER

if ($HASH_TEST_PATH) {
  if (!$HASH_PATH_FOLDER_TEST) {
    New-Item -Type Directory -Path $HASH_PATH_FOLDER
  }

  Copy-Item -Path "C:\HWID\*" -Destination "$HASH_PATH_FOLDER" -Recurse
  Set-Location "U:\SERVIDOR DEPLOYMENT\MDT01\HASH\$DATE\"
  $FILE_COPIED_SUCCESSFULY = Test-Path -Path $STRINGHWID

  if ($FILE_COPIED_SUCCESSFULY) {
    Write-Host -ForegroundColor Green "`nArquivo copiado com sucesso ao servidor."
  } else {
    Write-Host -ForegroundColor Red "Não foi possível copiar o arquivo ao servidor. Favor realizar a cópia manualmente ou a um pendrive."
    explorer "C:\HWID"
  }

} else {
  explorer "C:\HWID"
}

$net.RemoveNetworkDrive("U:");
Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000001

Write-Host -ForegroundColor Green "`nPressione qualquer tecla para finalizar...";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
