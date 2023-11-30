# [string[]]$FUNCIONARIOS = Get-Content -Path '.\Funcionarios_producao.txt'
net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

$FUNCIONARIOS_PROD = [IO.File]::ReadAllLines("\\172.18.3.4\d$\Logs\PRODUCAO\Lista_Funcionarios.txt")
$FUNCIONARIOS_RMA = [IO.File]::ReadAllLines("\\172.18.3.4\d$\Logs\RMA\Lista_Funcionarios.txt")

$FUNCIONARIOS_PROD
$FUNCIONARIOS_RMA

#TEST