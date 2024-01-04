# $WINDOWS_STATUS = Get-CimInstance SoftwareLicensingProduct -Filter "partialproductkey is not null" | Where-Object Name -like windows* | Select-Object -ExpandProperty LicenseStatus

# if ($WINDOWS_STATUS -eq 1) {
#   Write-Host -ForegroundColor Green "O windows está ativado."
# } else {
#   Write-Host -ForegroundColor Red "O Windows não está ativado."
# }

Import-CSV C:\Temp\01-2024.csv -Delimiter ';' | Select-Object * -unique | Export-Csv C:\Temp\01-2024.csv -NoTypeInformation