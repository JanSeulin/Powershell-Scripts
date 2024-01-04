net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

$INPUT_DATA = '\\172.18.3.4\d$\Logs\PRODUCAO\01-2024\01-2024.csv'

$INPUT_CSV = Import-Csv $INPUT_DATA

$PATRIMONIO = 301005
$SERIAL = "PE0AACEA"
$HORA = "14:58"

foreach ($LINE in $INPUT_CSV) {
  if (($LINE -like "*$patrimonio*") -and ($LINE -like "*$SERIAL*") -and ($LINE -like "*$HORA*")) {
    Write-Host "Value Found"
  }
 # Write-Host $LINE
}

# & "\\172.18.3.4\d$\Servidor Deployment\MDT01\Scripts\Gerar_Log\Gerar_Log_Autopilot.exe" 'TIMAC (AUTOPILOT)' 'TAGS-TIMAC-AGRO-BRASIL-COMMERCIAL'
<#

#>