PowerShell.exe -WindowStyle hidden {
	$LOOP = 'true'

  Start-Transcript -Path C:\Logs\Transcript\Log.txt -Force

  while ($LOOP) {
    try{
      net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

      $MONTH_YEAR = Get-Date -Format "MM-yyyy"
      $CORE_PATH_PROD = Test-Path "\\172.18.3.4\d$\Fonte_Dados\PRODUCAO\"
      $FOLDER_PATH_PROD = Test-Path "\\172.18.3.4\d$\Fonte_Dados\PRODUCAO\$MONTH_YEAR\"

      $CORE_PATH_RMA = Test-Path "\\172.18.3.4\d$\Fonte_Dados\RMA\"
      $FOLDER_PATH_RMA = Test-Path "\\172.18.3.4\d$\Fonte_Dados\RMA\$MONTH_YEAR\"

      if (!$CORE_PATH_PROD){
        New-Item -Type Directory -Path "\\172.18.3.4\d$\Fonte_Dados\PRODUCAO\"
      }

      if (!$FOLDER_PATH_PROD) {
        New-Item -Type Directory -Path "\\172.18.3.4\d$\Fonte_Dados\PRODUCAO\$MONTH_YEAR\"
      }

      if (!$CORE_PATH_RMA){
        New-Item -Type Directory -Path "\\172.18.3.4\d$\Fonte_Dados\RMA\"
      }

      if (!$FOLDER_PATH_RMA) {
        New-Item -Type Directory -Path "\\172.18.3.4\d$\Fonte_Dados\RMA\$MONTH_YEAR\"
      }

      Copy-Item -Path \\172.18.3.4\d$\Logs\PRODUCAO\$MONTH_YEAR\$MONTH_YEAR.csv	-Destination \\172.18.3.4\d$\Fonte_Dados\PRODUCAO\$MONTH_YEAR\$MONTH_YEAR.csv -Force -ErrorAction Stop
      Copy-Item -Path \\172.18.3.4\d$\Logs\PRODUCAO\$MONTH_YEAR\Autopilot.csv	-Destination \\172.18.3.4\d$\Fonte_Dados\PRODUCAO\$MONTH_YEAR\Autopilot.csv -Force -ErrorAction Stop
      Copy-Item -Path \\172.18.3.4\d$\Logs\RMA\$MONTH_YEAR\$MONTH_YEAR.csv	-Destination \\172.18.3.4\d$\Fonte_Dados\RMA\$MONTH_YEAR\$MONTH_YEAR.csv -Force -ErrorAction Stop

      $HOUR = Get-Date -Format HH:mm
      Write-Host -ForegroundColor Green "Comando executado com sucesso. Hora: $HOUR"
      Start-Sleep -Seconds 60

    } catch {
      Write-Host -ForeGroundColor Red "Algo deu errado, tentando novamente em alguns segundos..."
      Start-Sleep -Seconds 5
    }
  }
}






