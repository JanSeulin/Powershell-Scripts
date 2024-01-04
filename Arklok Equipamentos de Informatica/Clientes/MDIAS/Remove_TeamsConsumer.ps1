<#
    Script to check for an remove consumer Teams from Windows 11.
    This will also inject a registry key to make sure that Consumer Teams does not auto install later.
    The registry key is a protected key that is owned by Trusted Installer and cannot be modified by the SYSTEM account.
    This is the reason a scheduled task to be run as Trusted Installer will be used to modify this registry value.
#>

# Desabilitar ícone do Chat, se ativo
$TEST_CHAT_ICON = Test-Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsChat

if ($TEST_CHAT_ICON) {
  Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsChat -Name ChatIcon -Value 3
}

#region functions
function Test-RegistryValue {
  [CmdletBinding()]
  param (
      [Parameter(Mandatory = $true)]
      [string]$KeyPath,
      [Parameter(Mandatory = $true)]
      [string]$ValueName,
      [Parameter(Mandatory = $true)]
      [string]$ExpectedData
  )
  try {
      $actualData = Get-ItemPropertyValue -Path $KeyPath -Name $ValueName -ErrorAction Stop
      if ($actualData -eq $ExpectedData) {
          return $true
      }
      else {
          return $false
      }
  }
  catch {
      return $false
  }
}
#endregion functions

#region startlogging
$ScriptStartTime = Get-Date -Format "D-yyyyMMdd_T-HHmmss"
$LogFolder = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$logFile = "$LogFolder\ConsumerTeamsRemoval_Script_" + $ScriptStartTime + ".log"
Start-Transcript -Path $LogFile -Append
#endregion startlogging

#region RemoveTeams
$ConsumerTeams = Get-AppxPackage -Name MicrosoftTeams -AllUsers

if (!$ConsumerTeams) {
  Write-Host "Consumer teams app does not exist on the device."
  Write-Host "Next step is to add the value to the registry to prevent teams from auto installing in the future."
}


else {
  try {
      Write-Host "Consumer Teams removal will be attempted from $env:computername [$ConsumerTeams]"
      Get-AppxPackage -Name MicrosoftTeams -AllUsers -Verbose | Remove-AppxPackage -AllUsers -Verbose -ErrorAction Stop
      Write-Host "Microsoft Teams consumer app successfully removed"
      Write-Host "Next step is to add the value to the registry to prevent teams from auto installing in the future."
  }

  catch {
      Write-Error "Something went wrong removing Microsoft Teams consumer app---[$($_.Exception.Message)]"
      Write-Host "We will still attempt add the value to the registry to prevent teams from auto installing in the future."
  }
}
#endregion RemoveTeams

#region RegAutoInstall
$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Communications"
$registryValue = "ConfigureChatAutoInstall"
$ExpectedData = "0"

$RegistryData = Test-RegistryValue -KeyPath $registryPath -ValueName $registryValue -ExpectedData $ExpectedData

if ($RegistryData) {
  Write-Host "The registry value {$registryValue at the path $registryPath} exists and is set correctly."
}

else {
  try {

      Write-Host "Creating a scheduled task to run as Trusted Installer to inject the proper value into the registry to Prevent Teams AutoInstall..."

      $Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-executionpolicy bypass -command "reg.exe add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Communications /v ConfigureChatAutoInstall /t REG_DWORD /d 0 /f | Out-Host"'
      $Principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-544" #Warning: the admin Group name is localised

      Register-ScheduledTask -TaskName 'uninstallChat' -Action $action -Principal $Principal -Verbose

      $svc = New-Object -ComObject 'Schedule.Service'
      $svc.Connect()

      $user = 'NT SERVICE\TrustedInstaller'
      $folder = $svc.GetFolder('\')
      $task = $folder.GetTask('uninstallChat')

      #Start Task
      $task.RunEx($null, 0, 0, $user)
      Start-Sleep -Seconds 5 -Verbose

      #Kill Task
      $task.Stop(0)
      Start-Sleep -Seconds 2 -Verbose

      Write-Host "Task has run. Unregistering the task."

      #remove task From Task Scheduler
      Unregister-ScheduledTask -TaskName 'uninstallChat' -Confirm:$false -Verbose
      Start-Sleep -Seconds 2 -Verbose
  }
  catch {
      $_
  }

}
#region RegAutoInstall

#region CheckRegistry
#This will check if the registry value is correct. If not, it will fail so the script will attempt to run again later.
$RegistryData2 = Test-RegistryValue -KeyPath $registryPath -ValueName $registryValue -ExpectedData 0

if ($RegistryData2) {
  Write-Host "$registryValue is now set to the correct data."
}

else {
  Write-Warning "$registryValue is still not correct."
  Stop-Transcript
  Exit 1
}
#endregion CheckRegistry

Stop-Transcript