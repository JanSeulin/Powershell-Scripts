Get-CimInstance Win32_ComputerSystem | Select-Object -ExpandProperty SystemSKUNumber
# (Get-CimInstance -Namespace root\wmi -ClassName MS_SystemInformation).SystemSKU

# MUDAR LINGUAGEM DO WINDOWS
Set-WinUserLanguageList -LanguageList en-US -Force

# IDENTIFICAR TEMA DO WINDOWS
$WINDOWS_THEME = Get-ItemPropertyValue -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize AppsUseLightTheme
# Write-Host $WINDOWS_THEME

if ($WINDOWS_THEME -eq 0) {
  Write-Host "Your Windows is in dark mode"
} else {
  Write-Host "Your Windows is in light mode"
}

$TEST_CHAT_ICON = Test-Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsChat

if ($TEST_CHAT_ICON) {
  Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsChat -Name ChatIcon -Value 3
}

# CHECAR EM QUE AUTOPILOT MÁQUINA ESTÁ, SE EM ALGUM
$AUTOPILOT_TENANT_DOMAIN = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\Autopilot CloudAssignedTenantDomain
$AUTOPILOT_TENANT_ID = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\Autopilot CloudAssignedTenantId

$DOMAIN_SIMPLE = ""

if ($AUTOPILOT_TENANT_DOMAIN -like "mdias") {
  $DOMAIN_SIMPLE = "(MDIAS)"
} elseif ($AUTOPILOT_TENANT_DOMAIN -like "roullier") {
  $DOMAIN_SIMPLE = "(TIMAC)"
}

if ($AUTOPILOT_TENANT_DOMAIN) {
  Write-Host -ForegroundColor red "Atenção: Máquina já está registrada no seguinte domínio: $AUTOPILOT_TENANT_DOMAIN $DOMAIN_SIMPLE`n"
  Write-Host -ForegroundColor yellow "O processo de importação não será finalizado com sucesso se prosseguir."
  Write-Host -ForegroundColor yellow "Verificar se esse é o domínio desejado, caso contrário, entrar em contato com o time de Imagem para que seja solicitada a remoção da máquina do domínio atual.`n`n"

  $KEY_DOMAIN = ""
  while ($KEY_DOMAIN -ne 'E') {
    Write-Host -ForegroundColor Blue "Pressione E para prosseguir com a execução."
    $KEY_DOMAIN = [Console]::ReadKey($true).Key
  }
}

Write-Host $AUTOPILOT_TENANT_DOMAIN
Write-Host $AUTOPILOT_TENANT_Id