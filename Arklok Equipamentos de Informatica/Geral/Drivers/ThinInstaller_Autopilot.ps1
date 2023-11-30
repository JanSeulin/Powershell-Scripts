$COMPUTER_SYSTEM = Get-CimInstance Win32_ComputerSystem
$NOTEBOOK_MANUFACTURER = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Manufacturer

if ($NOTEBOOK_MANUFACTURER -like "*Lenovo*") {
  net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505
  & "\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\Scripts\ThinInstaller\lenovo_thininstaller.exe" /VERYSILENT

  Start-Sleep -Seconds 10

  & "C:\Program Files (x86)\Lenovo\ThinInstaller\ThinInstaller.exe" /CM -search A -action INSTALL -repository "\\172.18.3.4\d$\SERVIDOR DEPLOYMENT\MDT01\DRIVERS_TESTE\REPOSITORIO\LENOVO"-noicon -includerebootpackages 3 -noreboot
}