net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

$INPUT_DATA = '\\172.18.3.4\d$\Logs\PRODUCAO\11-2023\11-2023.csv'

$INPUT_CSV = Import-Csv $INPUT_DATA | Sort-Object * -Unique

$INPUT_CSV | Export-Csv "$INPUT_DATA-Test.csv" -NoTypeInformation



# & "\\172.18.3.4\d$\Servidor Deployment\MDT01\Scripts\Gerar_Log\Gerar_Log_Autopilot.exe" 'TIMAC (AUTOPILOT)' 'TAGS-TIMAC-AGRO-BRASIL-COMMERCIAL'