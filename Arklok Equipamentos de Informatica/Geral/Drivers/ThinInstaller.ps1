$DRIVERS_ERROR = (Get-CimInstance Win32_PNPEntity | Where-Object{[string]::IsNullOrEmpty($_.ClassGuid) } | Select-Object Name,Present,Status,DeviceID | Where-Object -Property Status -Contains 'Error' | Measure-Object).Count

if ($DRIVERS_ERROR -eq 0){
  Write-Host -ForegroundColor Green "Nenhum driver faltando, finalizando execução..."
  exit
}

$COMPUTER_SYSTEM = Get-CimInstance Win32_ComputerSystem
$NOTEBOOK_MANUFACTURER = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Manufacturer

function TestConnection {
  While (((Test-NetConnection www.google.com -Port 80 -InformationLevel "Detailed").TcpTestSucceeded) -eq $false) {
    Start-Sleep -Seconds 5
  }
}

function NetUseServer {
  net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505
}

if ($NOTEBOOK_MANUFACTURER -like "*Lenovo*") {
  NetUseServer
  Copy-Item '\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\Scripts\ThinInstaller\ThinInstaller' -Recurse -Destination "C:\Program Files (x86)\Lenovo\ThinInstaller" -Force

  & "C:\Program Files (x86)\Lenovo\ThinInstaller\ThinInstaller.exe" /CM -search R -action INSTALL -noicon -includerebootpackages 3 -noreboot -repository '\\172.18.3.4\D$\SERVIDOR DEPLOYMENT\MDT01\DRIVERS_TESTE\REPOSITORIO\LENOVO' | Out-Null

  TestConnection
} elseif ($NOTEBOOK_MANUFACTURER -like "*Dell*") {
  NetUseServer

  & "\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\Scripts\ThinInstaller\Dell_Update\Dell_CommandUpdate.EXE" /s | Out-Null

  & "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" /configure -silent -autoSuspendBitLocker=enable -userConsent=disable
  & "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" /scan -updateSeverity=recommended -outputLog=C:\dell\logs\scan.log
  & "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" /applyUpdates -reboot=disable -outputLog=C:\dell\logs\applyUpdates.log

  TestConnection
} elseif ($NOTEBOOK_MANUFACTURER -like "*SAMSUNG*") {
  NetUseServer
  $NOTEBOOK_MODEL = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Model
  $DRIVER_PATH_SAMSUNG = "\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\DRIVERS_TESTE\SAMSUNG\$NOTEBOOK_MODEL"
  Get-ChildItem $DRIVER_PATH_SAMSUNG -Recurse -Filter "*.inf" | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }

  TestConnection
} elseif ($NOTEBOOK_MANUFACTURER -like "*HP*") {
  NetUseServer

  & "\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\Scripts\ThinInstaller\HP_Update\hp-hpia-5.1.11.exe" /s /e /f C:\HPIA | Out-Null
  & "C:\HPIA\HPImageAssistant.exe" /Operation:Analyze /Category:drivers /selection:Critical /action:install /AutoCleanup /silent /reportFolder:c:\temp\HPIA\Report /softpaqdownloadfolder:c:\temp\HPIA\download | Out-Null

  TestConnection
}