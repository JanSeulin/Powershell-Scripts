Get-CimInstance Win32_ComputerSystem | Select-Object -ExpandProperty SystemSKUNumber
# (Get-CimInstance -Namespace root\wmi -ClassName MS_SystemInformation).SystemSKU


# IDENTIFICAR TEMA DO WINDOWS
$WINDOWS_THEME = Get-ItemPropertyValue -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize AppsUseLightTheme
Write-Host $WINDOWS_THEME

if ($WINDOWS_THEME -eq 0) {
  Write-Host "Your Windows is in dark mode"
} else {
  Write-Host "Your Windows is in light mode"
}

# CHECAR EM QUE AUTOPILOT MÁQUINA ESTÁ, SE EM ALGUM
$AUTOPILOT_TENANT_DOMAIN = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\Autopilot CloudAssignedTenantDomain
$AUTOPILOT_TENANT_ID = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\Autopilot CloudAssignedTenantId
Write-Host $AUTOPILOT_TENANT_DOMAIN
Write-Host $AUTOPILOT_TENANT_Id