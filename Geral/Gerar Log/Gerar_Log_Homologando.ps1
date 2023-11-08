Add-Type -AssemblyName PresentationFramework

# FUNÇÃO PARA MANIPULAÇÃO DA CAIXA VIA PRESENTATION FRAMEWORK - FONTE: https://smsagent.wordpress.com/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
Function New-WPFMessageBox {
  # Define Parameters
  [CmdletBinding()]
  Param
  (
      # The popup Content
      [Parameter(Mandatory=$True,Position=0)]
      [Object]$Content,

      # The window title
      [Parameter(Mandatory=$false,Position=1)]
      [string]$Title,

      # The buttons to add
      [Parameter(Mandatory=$false,Position=2)]
      [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
      [array]$ButtonType = 'OK',

      # The buttons to add
      [Parameter(Mandatory=$false,Position=3)]
      [array]$CustomButtons,

      # Content font size
      [Parameter(Mandatory=$false,Position=4)]
      [int]$ContentFontSize = 14,

      # Title font size
      [Parameter(Mandatory=$false,Position=5)]
      [int]$TitleFontSize = 14,

      # BorderThickness
      [Parameter(Mandatory=$false,Position=6)]
      [int]$BorderThickness = 0,

      # CornerRadius
      [Parameter(Mandatory=$false,Position=7)]
      [int]$CornerRadius = 8,

      # ShadowDepth
      [Parameter(Mandatory=$false,Position=8)]
      [int]$ShadowDepth = 3,

      # BlurRadius
      [Parameter(Mandatory=$false,Position=9)]
      [int]$BlurRadius = 20,

      # WindowHost
      [Parameter(Mandatory=$false,Position=10)]
      [object]$WindowHost,

      # Timeout in seconds,
      [Parameter(Mandatory=$false,Position=11)]
      [int]$Timeout,

      # Code for Window Loaded event,
      [Parameter(Mandatory=$false,Position=12)]
      [scriptblock]$OnLoaded,

      # Code for Window Closed event,
      [Parameter(Mandatory=$false,Position=13)]
      [scriptblock]$OnClosed

  )

  # Dynamically Populated parameters
  DynamicParam {

      # Add assemblies for use in PS Console
      Add-Type -AssemblyName System.Drawing, PresentationCore

      # ContentBackground
      $ContentBackground = 'ContentBackground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ContentBackground = "White"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)


      # FontFamily
      $FontFamily = 'FontFamily'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Drawing.FontFamily]::Families.Name | Select -Skip 1
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
      $PSBoundParameters.FontFamily = "Segoe UI"

      # TitleFontWeight
      $TitleFontWeight = 'TitleFontWeight'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.TitleFontWeight = "Normal"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

      # ContentFontWeight
      $ContentFontWeight = 'ContentFontWeight'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ContentFontWeight = "Normal"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)


      # ContentTextForeground
      $ContentTextForeground = 'ContentTextForeground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ContentTextForeground = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

      # TitleTextForeground
      $TitleTextForeground = 'TitleTextForeground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.TitleTextForeground = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

      # BorderBrush
      $BorderBrush = 'BorderBrush'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.BorderBrush = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


      # TitleBackground
      $TitleBackground = 'TitleBackground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.TitleBackground = "White"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

      # ButtonTextForeground
      $ButtonTextForeground = 'ButtonTextForeground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ButtonTextForeground = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

      # Sound
      $Sound = 'Sound'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      #$ParameterAttribute.Position = 14
      $AttributeCollection.Add($ParameterAttribute)
      $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

      return $RuntimeParameterDictionary
  }

  Begin {
      Add-Type -AssemblyName PresentationFramework
  }
#WindowStartupLocation="Manual" Top="0"
  Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1" Topmost="True" Top="0"
      WindowStartupLocation="CenterScreen">
  <Window.Resources>
      <Style TargetType="{x:Type Button}">
          <Setter Property="Template">
              <Setter.Value>
                  <ControlTemplate TargetType="Button">
                      <Border>
                          <Grid Background="{TemplateBinding Background}">
                              <ContentPresenter />
                          </Grid>
                      </Border>
                  </ControlTemplate>
              </Setter.Value>
          </Setter>
      </Style>
  </Window.Resources>
  <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
      <Border.Effect>
          <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.3" />
      </Border.Effect>
      <Border.Triggers>
          <EventTrigger RoutedEvent="Window.Loaded">
              <BeginStoryboard>
                  <Storyboard>
                      <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                      <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                  </Storyboard>
              </BeginStoryboard>
          </EventTrigger>
      </Border.Triggers>
      <Grid >
          <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
          <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
              <Grid.OpacityMask>
                  <VisualBrush Visual="{Binding ElementName=Mask}"/>
              </Grid.OpacityMask>
              <StackPanel Name="StackPanel" >
                  <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                  <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                  </DockPanel>
                  <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                  </DockPanel>
              </StackPanel>
          </Grid>
      </Grid>
  </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="24" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="24" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

  # Load the window from XAML
  $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

  # Custom function to add a button
  Function Add-Button {
      Param($Content)
      $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
      $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
      $ButtonText.Text = "$Content"
      $Button.Content = $ButtonText
      $Button.Add_MouseEnter({
          $This.Content.FontSize = "25"
      })
      $Button.Add_MouseLeave({
          $This.Content.FontSize = "24"
      })
      $Button.Add_Click({
          # New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
          New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Scope Global -Force
          $Window.Close()
      })
      $Window.FindName('ButtonHost').AddChild($Button)
  }

  # Add buttons
  If ($ButtonType -eq "OK")
  {
      Add-Button -Content "OK"
  }

  If ($ButtonType -eq "OK-Cancel")
  {
      Add-Button -Content "OK"
      Add-Button -Content "Cancel"
  }

  If ($ButtonType -eq "Abort-Retry-Ignore")
  {
      Add-Button -Content "Abort"
      Add-Button -Content "Retry"
      Add-Button -Content "Ignore"
  }

  If ($ButtonType -eq "Yes-No-Cancel")
  {
      Add-Button -Content "Yes"
      Add-Button -Content "No"
      Add-Button -Content "Cancel"
  }

  If ($ButtonType -eq "Yes-No")
  {
      Add-Button -Content "Yes"
      Add-Button -Content "No"
  }

  If ($ButtonType -eq "Retry-Cancel")
  {
      Add-Button -Content "Retry"
      Add-Button -Content "Cancel"
  }

  If ($ButtonType -eq "Cancel-TryAgain-Continue")
  {
      Add-Button -Content "Cancel"
      Add-Button -Content "TryAgain"
      Add-Button -Content "Continue"
  }

  If ($ButtonType -eq "None" -and $CustomButtons)
  {
      Foreach ($CustomButton in $CustomButtons)
      {
          Add-Button -Content "$CustomButton"
      }
  }

  # Remove the title bar if no title is provided
  If ($Title -eq "")
  {
      $TitleBar = $Window.FindName('TitleBar')
      $Window.FindName('StackPanel').Children.Remove($TitleBar)
  }

  # Add the Content
  If ($Content -is [String])
  {
      # Replace double quotes with single to avoid quote issues in strings
      If ($Content -match '"')
      {
          $Content = $Content.Replace('"',"'")
      }

      # Use a text box for a string value...
      $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
      $Window.FindName('ContentHost').AddChild($ContentTextBox)
  }
  Else
  {
      # ...or add a WPF element as a child
      Try
      {
          $Window.FindName('ContentHost').AddChild($Content)
      }
      Catch
      {
          $_
      }
  }

  # Enable window to move when dragged
  $Window.FindName('Grid').Add_MouseLeftButtonDown({
      $Window.DragMove()
  })

  # Activate the window on loading
  If ($OnLoaded)
  {
      $Window.Add_Loaded({
          $This.Activate()
          Invoke-Command $OnLoaded
      })
  }
  Else
  {
      $Window.Add_Loaded({
          $This.Activate()
      })
  }


  # Stop the dispatcher timer if exists
  If ($OnClosed)
  {
      $Window.Add_Closed({
          If ($DispatcherTimer)
          {
              $DispatcherTimer.Stop()
          }
          Invoke-Command $OnClosed
      })
  }
  Else
  {
      $Window.Add_Closed({
          If ($DispatcherTimer)
          {
              $DispatcherTimer.Stop()
          }
      })
  }


  # If a window host is provided assign it as the owner
  If ($WindowHost)
  {
      $Window.Owner = $WindowHost
      $Window.WindowStartupLocation = "CenterOwner"
  }

  # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
  If ($Timeout)
  {
      $Stopwatch = New-object System.Diagnostics.Stopwatch
      $TimerCode = {
          If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
          {
              $Stopwatch.Stop()
              $Window.Close()
          }
      }
      $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
      $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
      $DispatcherTimer.Add_Tick($TimerCode)
      $Stopwatch.Start()
      $DispatcherTimer.Start()
  }

  # Play a sound
  If ($($PSBoundParameters.Sound))
  {
      $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
      $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
      $SoundPlayer.Add_LoadCompleted({
          $This.Play()
          $This.Dispose()
      })
      $SoundPlayer.LoadAsync()
  }

  # Display the window
  $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

  }
}

$param1 = $args[0]
Write-Host $param1

net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

#################### PATRIMONIO ####################
$PATRIMONIO_LOOP = "true"

While ($PATRIMONIO_LOOP -eq "true") {
  # Create a textblock
  $TextBlock_Patrimonio = New-Object System.Windows.Controls.TextBlock
  $TextBlock_Patrimonio.Text = "Digite o Patrimônio da Máquina"
  $TextBlock_Patrimonio.Margin = 10
  $TextBlock_Patrimonio.FontSize = 16

  $TextBox_Patrimonio = New-Object System.Windows.Controls.TextBox
  $TextBox_Patrimonio.Text = ""
  $TextBox_Patrimonio.Margin = "10,0,10,0"
  $TextBox_Patrimonio.FontSize = 16
  $TextBox_Patrimonio.BorderThickness = 2

  $StackPanel_Patrimonio = New-Object System.Windows.Controls.StackPanel
  $StackPanel_Patrimonio.AddChild($TextBlock_Patrimonio)
  $StackPanel_Patrimonio.AddChild($TextBox_Patrimonio)

  $Params_Patrimonio=@{
    Content=$StackPanel_Patrimonio
    Title="Patrimônio"
    TitleBackground="DarkRed"
    TitleFontSize=20
    TitleFontWeight='Bold'
    TitleTextForeground='White'
    ButtonType='None'
    ButtonTextForeground="DarkRed"
    CustomButtons="OK"
    BorderThickness=2
    ShadowDepth=4
    ContentFontSize=1
  }

  New-WPFMessageBox @Params_Patrimonio

  if ($WPFMessageBoxOutput -eq "OK") {
    $PATRIMONIO = $TextBox_Patrimonio.text

    Write-Host "PATRIMONIO: $PATRIMONIO"

    if (($PATRIMONIO.Length -lt 4) -or ($PATRIMONIO.Length -gt 6) -or ($PATRIMONIO -match "[^0-9]")) {
      $Params_Patrimonio_Erro=@{
        Content="Valor inválido. O patrimônio deve ter entre 4 e 6 caracteres numéricos."
        Title="Erro"
        TitleBackground="DarkRed"
        TitleFontSize=20
        TitleFontWeight='Bold'
        TitleTextForeground='White'
        ButtonType='None'
        ButtonTextForeground="DarkRed"
        CustomButtons="OK"
        BorderThickness=2
        ShadowDepth=4
        ContentFontSize=16
      }

      New-WPFMessageBox @Params_Patrimonio_Erro
    } else {
      $PATRIMONIO_LOOP = "false"
    }
  }
}


$LOOP_SETOR = "true";

while  ($LOOP_SETOR) {
  ################ SETOR ###############
  $TextBlock_Setor = New-Object System.Windows.Controls.TextBlock
  $TextBlock_Setor.Text = "Atenção: não é possível alterar o setor e colaborador após confirmação."
  $TextBlock_Setor.Margin = 10
  $TextBlock_Setor.FontSize = 16

  $StackPanel_SETOR = New-Object System.Windows.Controls.StackPanel
  # $StackPanel_SETOR.AddChild($IMAGE)
  $StackPanel_SETOR.AddChild($TextBlock_Setor)

  $Params_Setor=@{
    Content=$StackPanel_SETOR
    Title="Selecione o Setor"
    TitleBackground="DarkRed"
    TitleFontSize=20
    TitleFontWeight='Bold'
    TitleTextForeground='White'
    ButtonType='None'
    ButtonTextForeground="DarkRed"
    CustomButtons="PRODUCAO","RMA"
    BorderThickness=2
    ShadowDepth=4
    ContentFontSize=1
  }

  $SETOR = ""

  New-WPFMessageBox @Params_Setor

  if ($WPFMessageBoxOutput -eq "Producao") {
    $SETOR = "PRODUCAO"
  } elseif ($WPFMessageBoxOutput -eq "RMA") {
    $SETOR = "RMA"
  }

  [string[]]$LISTA_PRODUCAO = Get-Content -Path '\\172.18.3.4\d$\Logs\PRODUCAO\Lista_Funcionarios.txt'
  [string[]]$LISTA_RMA = Get-Content -Path '\\172.18.3.4\d$\Logs\RMA\Lista_Funcionarios.txt'

  #################### COLABORADOR ####################
  $StackPanel_COLABORADOR = New-Object System.Windows.Controls.StackPanel
  $ComboBox = New-Object System.Windows.Controls.ComboBox
  $ComboBox.Margin = "10,10,10,0"
  $ComboBox.Background = "White"
  $ComboBox.FontSize = 16

  if ($SETOR -eq "PRODUCAO") {
    $ComboBox.ItemsSource = $LISTA_PRODUCAO
  } else {
    $ComboBox.ItemsSource = $LISTA_RMA
  }

  $RANDOM_INT = Get-Random -Maximum ($ComboBox.ItemsSource | Measure-Object).Count
  $ComboBox.SelectedIndex = $RANDOM_INT

  $TextBlock = New-Object System.Windows.Controls.TextBlock
  $TextBlock.Text = "Selecione seu nome na lista abaixo:"
  $TextBlock.Margin = 10
  $TextBlock.FontSize = 16

  $TextBlock, $ComboBox | ForEach-Object {
    $StackPanel_COLABORADOR.AddChild($PSItem)
  }

  $Params_COLABORADOR=@{
    Content=$StackPanel_COLABORADOR
    Title="Colaborador"
    TitleBackground="DarkRed"
    TitleFontSize=20
    TitleFontWeight='Bold'
    TitleTextForeground='White'
    ButtonType='None'
    ButtonTextForeground="DarkRed"
    CustomButtons="Confirmar","Voltar"
    BorderThickness=2
    ShadowDepth=4
    ContentFontSize=1
  }

  $ComboBox | Select-Object *

  New-WPFMessageBox @Params_COLABORADOR

  if ($WPFMessageBoxOutput -eq "Confirmar") {
    $COLABORADOR = $ComboBox.SelectedValue
    break;
  } elseif ($WPFMessageBoxOutput -eq "Voltar") {
    Continue;
  }

  Write-Host "`n"
}


# net use \\172.18.3.4\d$ /user:arkserv\jan Lucy@505

$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
$CLIENTE = $tsenv.Value("TaskSequenceName")

# $DAY = Get-Date -Format "dd - dddd"
$DAY_MONTH = Get-Date -Format "dd - MM"
$MONTH_YEAR = Get-Date -Format "MM-yyyy"
$HOUR_MINUTE = Get-Date -Format "HH:mm"

# "\\172.18.3.4\d$\Logs\$SETOR\$MONTH_YEAR"
$TEST_PATH = Test-Path "\\172.18.3.4\d$\Logs\$SETOR\$MONTH_YEAR"
$TEST_FILE_PATH = Test-Path "\\172.18.3.4\d$\Logs\$SETOR\$MONTH_YEAR\$MONTH_YEAR.csv"

if (!$TEST_PATH) {
  New-Item -Type Directory -Path "\\172.18.3.4\d$\Logs\$SETOR\$MONTH_YEAR"
}

if (!$TEST_FILE_PATH) {
  Set-Location -Path "\\172.18.3.4\d$\Logs\$SETOR\$MONTH_YEAR"
  New-Item ".\$MONTH_YEAR.csv"
}

$FULL_PATH = "\\172.18.3.4\d$\Logs\$SETOR\$MONTH_YEAR\$MONTH_YEAR.csv"
$LINES = (Get-Content $FULL_PATH | Measure-Object).Count

$COMPUTER_SYSTEM = Get-CimInstance Win32_ComputerSystem
$NOTEBOOK_MANUFACTURER = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Manufacturer
$NOTEBOOK_MODEL = $COMPUTER_SYSTEM | Select-Object -ExpandProperty SystemFamily

if ($NOTEBOOK_MANUFACTURER -like "*Dell*") {
  $NOTEBOOK_MANUFACTURER = "Dell"
  $NOTEBOOK_MODEL = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Model
} elseif ($NOTEBOOK_MANUFACTURER -like "*LENOVO*") {
  $NOTEBOOK_MANUFACTURER = "Lenovo"
} elseif ($NOTEBOOK_MANUFACTURER -like "*SAMSUNG*") {
  $NOTEBOOK_MANUFACTURER = "Samsung"
} elseif ($NOTEBOOK_MANUFACTURER -like "*HP*") {
  $NOTEBOOK_MANUFACTURER = ""
  $NOTEBOOK_MODEL = $COMPUTER_SYSTEM | Select-Object -ExpandProperty Model
}

$NOTEBOOK_STRING = $NOTEBOOK_MANUFACTURER + " " + $NOTEBOOK_MODEL

$MEMORY = Get-CimInstance Win32_PhysicalMemory
$MEMORY_SIZE = ($MEMORY | Measure-Object -Property capacity -Sum).sum /1gb
$MEMORY_STRING = [string]($MEMORY_SIZE) + " GB"

$STORAGE_BASE = Get-CimInstance Win32_DiskDrive
$STORAGE_STRING = [string][math]::Round((($STORAGE_BASE | Where-Object -Property MediaType -ne "Removable Media" | Measure-Object -Property Size -sum).sum)/1GB, 2) + " GB"

$SERIAL = Get-CimInstance Win32_Bios | Select-Object -ExpandProperty SerialNumber

$PARAMS_ERROR = @{
  Title="Erro"
  Content="Não foi possível salvar as informações, tentando novamente em alguns segundos.."
  TitleBackground="DarkRed"
  TitleFontSize=20
  TitleFontWeight='Bold'
  TitleTextForeground='White'
  ButtonType='None'
  ButtonTextForeground="DarkRed"
  # CustomButtons="Confirmar"
  BorderThickness=2
  ShadowDepth=4
  ContentFontSize=16
  Timeout=5
}

$PARAMS_SUCCESS = @{
  Title="Sucesso"
  Content="Informações salvas com sucesso, finalizando..."
  TitleBackground="DarkRed"
  TitleFontSize=20
  TitleFontWeight='Bold'
  TitleTextForeground='White'
  ButtonType='None'
  ButtonTextForeground="DarkRed"
  # CustomButtons="Confirmar"
  BorderThickness=2
  ShadowDepth=4
  ContentFontSize=16
  Timeout=6
}

$SAVING_LOOP = 'true';

while ($SAVING_LOOP) {
  try {
    if ($LINES -eq 0) {
      "PATRIMONIO;CLIENTE;SERIAL;MODELO;MEMORIA;ARMAZENAMENTO;PRODUZIDO POR;SETOR;HORA;DIA" | Add-Content $FULL_PATH -ErrorAction Stop
    }

    "$PATRIMONIO;$CLIENTE;$SERIAL;$NOTEBOOK_STRING;$MEMORY_STRING;$STORAGE_STRING;$COLABORADOR;$SETOR;$HOUR_MINUTE;$DAY_MONTH" | Add-Content $FULL_PATH -ErrorAction Stop

    New-WPFMessageBox @PARAMS_SUCCESS
    exit
  }
  catch {
    <#Do this if a terminating exception happens#>
    $RANDOM_WAIT = Get-Random -Minimum 1 -Maximum 5
    Start-Sleep -Seconds $RANDOM_WAIT
    New-WPFMessageBox @PARAMS_ERROR
  }
}


