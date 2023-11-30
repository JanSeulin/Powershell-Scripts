$COMPUTER_SYSTEM = Get-CimInstance Win32_ComputerSystem
$NOTEBOOK_MANUFACTURER = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Manufacturer

if ($NOTEBOOK_MANUFACTURER -like "*LENOVO*") {
	Install-Module -Name 'LSUClient'
	$updates = Get-LSUpdate | Where-Object { $_.Installer.Unattended }
	$updates | Save-LSUpdate -Verbose
	$updates | Install-LSUpdate -Verbose
}