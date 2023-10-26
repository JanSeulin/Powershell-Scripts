# Verificar se há erro de driver, se sim, rodar script, caso contrário, pular execução
$DRIVERS_ERROR = (Get-CimInstance Win32_PNPEntity | Where-Object{[string]::IsNullOrEmpty($_.ClassGuid) } | Select-Object Name,Present,Status,DeviceID | Where-Object -Property Status -Contains 'Error' | Measure-Object).Count
#$DRIVERS_ERROR_MEASURE = $DRIVERS_ERROR | Measure-Object
#$DRIVERS_ERROR_COUNT = $DRIVERS_ERROR_MEASURE.Count

if ($DRIVERS_ERROR -eq 0){
	exit
}

$CHECK1 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' | Select-Object -ExpandProperty QuickEdit
$CHECK2 = Get-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' | Select-Object -ExpandProperty QuickEdit

if (($CHECK1 -eq 1) -or ($CHECK2 -eq 1)) {
  Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000000
  Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000000

  Start-Process -FilePath .\Drivers_MTM.exe
  exit
}

# PEGAR MARCA DA MÁQUINA
$MACHINE_MANUFACTURER = Get-ComputerInfo -Property "CsManufacturer" | Select-Object -ExpandProperty "CsManufacturer"
# PEGAR MTM DA MÁQUINA
$MACHINE_MTM = Get-ComputerInfo -Property "CsModel" | Select-Object -ExpandProperty "CsModel"
Write-Output $MACHINE_MTM
# PEGAR MODELO DA MÁQUINA
$MACHINE_FAMILY = Get-ComputerInfo -Property "CsSystemFamily" | Select-Object -ExpandProperty "CsSystemFamily"

# PEGAR VERSÃO DO WINDOWS E DECLARAR VARIÁVEL DE ACORDO
$WINDOWS_VER = Get-ComputerInfo -Property "OSName" | Select-Object -ExpandProperty "OSName"

if ($WINDOWS_VER -like "*Windows 10*"){
	$WINDOWS_VER = "Windows 10"
} else {
	$WINDOWS_VER = "Windows 11"
}

# CONSTRUÇÃO DO CAMINHO DOS DRIVERS DE ACORDO COM MARCA, MODELO E/OU MTM DA MÁQUINA E VERSÃO DO WINDOWS
switch -Wildcard ($MACHINE_MANUFACTURER) {
  "*Lenovo*" {
      $DRIVER_PATH = "..\..\DRIVERS_TESTE\LENOVO\" + $MACHINE_MTM
      $DRIVER_PATH_WINDOWS = $DRIVER_PATH + "\" + $WINDOWS_VER
      Write-Output $DRIVER_PATH
      Write-Output $DRIVER_PATH_WINDOWS
   }
   "*Dell*" {
      $DRIVER_PATH = "..\..\DRIVERS_TESTE\DELL\" + $MACHINE_MTM
      $DRIVER_PATH_WINDOWS = $DRIVER_PATH + "\" + $WINDOWS_VER
      Write-Output $DRIVER_PATH
      Write-Output $DRIVER_PATH_WINDOWS
   }
   "*Samsung*" {
      $DRIVER_PATH = "..\..\DRIVERS_TESTE\SAMSUNG\" + $MACHINE_FAMILY
      $DRIVER_PATH_WINDOWS = $DRIVER_PATH + "\" + $WINDOWS_VER
      Write-Output $DRIVER_PATH
      Write-Output $DRIVER_PATH_WINDOWS
   }
  Default {}
}

# TESTA SE OS CAMINHOS SÃO VÁLIDOS PARA TOMADA DE DECISÃO
$PATH_EXIST = Test-Path -Path $DRIVER_PATH
$PATH_EXIST_WINDOWS = Test-Path -Path $DRIVER_PATH_WINDOWS

if ($PATH_EXIST -eq "true") {
	#Write-Output "MTM PATH EXISTS"
	if ($PATH_EXIST_WINDOWS -eq "true") {
		#Write-Output "WINDOWS FOLDER EXISTS"
		Get-ChildItem $DRIVER_PATH_WINDOWS -Recurse -Filter "*.inf" |
		ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }
	} else {
		Get-ChildItem $DRIVER_PATH -Recurse -Filter "*.inf" |
		ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }
	}
} else {
	exit
}

Set-ItemProperty -Path 'HKCU:\Console' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_Powershell.exe' -Name 'QuickEdit' -Value 0x00000001
Set-ItemProperty -Path 'HKCU:\Console\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe' -Name 'QuickEdit' -Value 0x00000001

# # GERAR CAMINHO DA PASTA DE DRIVERS DE ACORDO COM MTM e VER Windows
# $DRIVER_PATH = "..\..\DRIVERS_TESTE\" + $MACHINE_MTM
# $DRIVER_PATH_WINDOWS = "..\..\DRIVERS_TESTE\" + $MACHINE_MTM + "\" + $WINDOWS_VER
# Write-Output $DRIVER_PATH
# Write-Output $DRIVER_PATH_WINDOWS

#Write-Output $PATH_EXIST
#Write-Host -NoNewLine 'Pressione qualquer tecla para finalizar...';
#$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

###### LINHAS PARA TESTE #####
#$Executable_Path = $DRIVER_PATH + "\" + "Teste.exe"
# Write-Output $DRIVER_PATH
# Write-Host -NoNewLine 'Pressione qualquer tecla para finalizar...';
# $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

#CHECAR SE CAMINHO JÁ EXISTE E SE SIM, REALIZAR A INSTALAÇÃO DOS DRIVERS. CASO CONTRÁRIO, SIMPLESMENTE ENCERRAR
#$PATH_EXIST = Test-Path -Path $DRIVER_PATH

#if ($PATH_EXIST -eq "true"){
#	Get-ChildItem $DRIVER_PATH -Recurse -Filter "*.inf" |
#	ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }
#} else {
#	exit
#}