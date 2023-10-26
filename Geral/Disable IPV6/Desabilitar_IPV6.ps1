Get-NetAdapterBinding | Where-Object ComponentID -EQ 'ms_tcpip6'
Disable-NetAdapterBinding -Name 'Ethernet' -ComponentID 'ms_tcpip6'
Disable-NetAdapterBinding -Name 'Wi-Fi' -ComponentID 'ms_tcpip6'
Get-NetAdapterBinding -Name 'Ethernet' -ComponentID 'ms_tcpip6'

Write-Host -NoNewLine 'Pressione qualquer tecla para finalizar...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');