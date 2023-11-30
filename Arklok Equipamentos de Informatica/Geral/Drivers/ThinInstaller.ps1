$COMPUTER_SYSTEM = Get-CimInstance Win32_ComputerSystem
$NOTEBOOK_MANUFACTURER = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Manufacturer

if ($NOTEBOOK_MANUFACTURER -like "*Lenovo*") {
  net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505
  Copy-Item '\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\Scripts\ThinInstaller\ThinInstaller' -Recurse -Destination "C:\Program Files (x86)\Lenovo\" -Force
  # & "Z:\Scripts\ThinInstaller\lenovo_thininstaller.exe" /VERYSILENT

  Start-Sleep -Seconds 5

  & "C:\Program Files (x86)\Lenovo\ThinInstaller\ThinInstaller.exe" /CM -search A -action INSTALL -noicon -includerebootpackages 3 -noreboot -repository '\\172.18.3.4\D$\SERVIDOR DEPLOYMENT\MDT01\DRIVERS_TESTE\REPOSITORIO\LENOVO'
}