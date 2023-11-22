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
      x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1" Topmost="True" Top="-10">
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
              <StackPanel Name="StackPanel">
                  <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                  <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                  </DockPanel>
                  <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center">
                  </DockPanel>
              </StackPanel>
          </Grid>
      </Grid>
  </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
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
          $This.Content.FontSize = "17"
      })
      $Button.Add_MouseLeave({
          $This.Content.FontSize = "16"
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





$LOOP = 'true'

while ($LOOP) {
  ############################ HARDWARE INFO ########################
  ### MODELO NOTEBOOK ###
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

  ### CPU ###
  $CPU_BASE = Get-CimInstance -ClassName Win32_Processor
  $CPU = $CPU_BASE | Select-Object -ExpandProperty Name
  $CPU_CORES = $CPU_BASE | Select-Object -ExpandProperty NumberOfCores
  $CPU_THREADS = $CPU_BASE | Select-Object -ExpandProperty NumberOfLogicalProcessors
  $CPU_SPEED = $CPU_BASE | Select-Object -ExpandProperty MaxClockSpeed

  ### MEMORIA ####
  $MEMORY = Get-CimInstance Win32_PhysicalMemory
  $MEMORY_SIZE = ($MEMORY | Measure-Object -Property capacity -Sum).sum /1gb
  $MEMORY_SLOTS = ($MEMORY | Select-Object BankLabel | Measure-Object).Count
  $TOTAL_NUMBER_OF_SLOTS = Get-CimInstance Win32_PhysicalMemoryArray | Select-Object -ExpandProperty MemoryDevices

  $SLOT_SIZE_STRING = ""
  $COUNT = 1

  foreach ($MEMORY_ITEM in $MEMORY) {
    $SIZE = [string]($MEMORY_ITEM | Measure-Object -Property capacity -Sum).sum /1gb

    if ($MEMORY_ITEM -eq $MEMORY[-1]) {
      $SLOT_SIZE_STRING += "Slot $COUNT`: $SIZE GB"
    } else {
      $SLOT_SIZE_STRING += "Slot $COUNT`: $SIZE GB - "
    }
    $COUNT++
  }

  $MEMORY_VOLTAGE = ($MEMORY  | Select-Object -ExpandProperty ConfiguredVoltage)
  $MEMORY_VOLTAGE = ($MEMORY_VOLTAGE -split '\n')[0]

  $MEMORY_SPEED = $MEMORY  | Select-Object -ExpandProperty Speed
  $MEMORY_SPEED = ($MEMORY_SPEED -split '\n')[0]

  # ARMAZENAMENTO FÓRMULA ATUAL - PERMITE MÚLTIPLOS HDS SEREM EXIBIDOS CORRETAMENTE
  $STORAGE_BASE = Get-CimInstance Win32_DiskDrive

  # ADICIONA TIPO DE MEMÓRIA DE ACORDO COM VOLTAGEM
  switch ($MEMORY_VOLTAGE){
    1100 { $MEMORY_TYPE = "DDR5"}
    1200 { $MEMORY_TYPE = "DDR4" }
    1500 { $MEMORY_TYPE = "DDR3" }
    1800 { $MEMORY_TYPE = "DDR2" }
    2500 { $MEMORY_TYPE = "DDR1" }
  }

  #### IMAGEM ARKLOK ####
  $SOURCE_IMG = "Z:\Scripts\Info_Sistema\Logo\Arklok_Slogan.png"
  $SOURCE_IMG_DEV="C:\Images\Arklok_Slogan.png"
  $IMAGE = New-Object System.Windows.Controls.Image
  $IMAGE.Source = $SOURCE_IMG
  $Image.Height = 100
  $Image.Width = 450
  $Image.Margin = "0,25,0,0"


  # NOTEBOOK MODELO
  $NOTEBOOK_BLOCK = New-Object System.Windows.Controls.TextBlock
  $NOTEBOOK_BLOCK.Text = "$NOTEBOOK_MANUFACTURER $NOTEBOOK_MODEL"
  $NOTEBOOK_BLOCK.margin = "0,5,0,5"
  $NOTEBOOK_BLOCK.FontSize = "17"
  $NOTEBOOK_BLOCK.HorizontalAlignment = "center"
  $NOTEBOOK_BLOCK.FontWeight = "Bold"
  $NOTEBOOK_BLOCK.Foreground = "DarkRed"

  # CPU TÍTULO
  $CPU_BLOCK_LABEL = New-Object System.Windows.Controls.TextBlock
  $CPU_BLOCK_LABEL.Text = "PROCESSADOR"
  $CPU_BLOCK_LABEL.margin = "20,5,5,5"
  $CPU_BLOCK_LABEL.FontSize = "16"
  $CPU_BLOCK_LABEL.FontWeight = "Bold"

  # CPU MODELO
  $CPU_BLOCK = New-Object System.Windows.Controls.TextBlock
  $CPU_BLOCK.Text = "Modelo: $CPU"
  $CPU_BLOCK.margin = "30,-5,15,0"
  $CPU_BLOCK.FontSize = "16"

  # CPU NÚCLEOS
  $CPU_CORES_BLOCK = New-Object System.Windows.Controls.TextBlock
  $CPU_CORES_BLOCK.Text = "Nucleos: $CPU_CORES, Threads: $CPU_THREADS"
  $CPU_CORES_BLOCK.margin = "30,0,15,0"
  $CPU_CORES_BLOCK.FontSize = "16"

  # CPU FREQUENCIA
  $CPU_SPEED_BLOCK = New-Object System.Windows.Controls.TextBlock
  $CPU_SPEED_BLOCK.Text = "Frequencia: $CPU_SPEED mhz"
  $CPU_SPEED_BLOCK.margin = "30,0,15,0"
  $CPU_SPEED_BLOCK.FontSize = "16"

  # MEMÓRIA TITULO
  $MEMORY_BLOCK_LABEL = New-Object System.Windows.Controls.TextBlock
  $MEMORY_BLOCK_LABEL.Text = "MEMORIA RAM"
  $MEMORY_BLOCK_LABEL.margin = "20,5,5,5"
  $MEMORY_BLOCK_LABEL.FontSize = "16"
  $MEMORY_BLOCK_LABEL.FontWeight = "Bold"

  # MEMÓRIA TAMANHO E TIPO
  $MEMORY_BLOCK = New-Object System.Windows.Controls.TextBlock
  $MEMORY_BLOCK.Text = "Total: $MEMORY_SIZE GB $MEMORY_TYPE"
  $MEMORY_BLOCK.margin = "30,-5,15,0"
  $MEMORY_BLOCK.FontSize = "16"

  # MEMÓRIA SLOTS
  $MEMORY_SLOTS_BLOCK = New-Object System.Windows.Controls.TextBlock

  $MEMORY_SLOTS_BLOCK.margin = "30,0,15,0"
  $MEMORY_SLOTS_BLOCK.FontSize = "16"

  $MEMORY_SLOTS_BLOCK = New-Object System.Windows.Controls.TextBlock
  $MEMORY_SLOTS_BLOCK.Text = "Slots em uso: $MEMORY_SLOTS de $TOTAL_NUMBER_OF_SLOTS"
  if ($NOTEBOOK_MANUFACTURER -eq "Samsung") {
    $MEMORY_SLOTS_BLOCK.Text = "Slots em uso: $MEMORY_SLOTS"
  }
  $MEMORY_SLOTS_BLOCK.margin = "30,0,15,0"
  $MEMORY_SLOTS_BLOCK.FontSize = "16"

  # MEMÓRIA FREQUENCIA
  $MEMORY_SPEED_BLOCK = New-Object System.Windows.Controls.TextBlock
  $MEMORY_SPEED_BLOCK.Text = "Frequencia: $MEMORY_SPEED mhz"
  $MEMORY_SPEED_BLOCK.margin = "30,0,15,0"
  $MEMORY_SPEED_BLOCK.FontSize = "16"

  # ARMAZENAMENTO TÍTULO
  $STORAGE_BLOCK_LABEL = New-Object System.Windows.Controls.TextBlock
  $STORAGE_BLOCK_LABEL.Text = "ARMAZENAMENTO"
  $STORAGE_BLOCK_LABEL.margin = "20,5,5,0"
  $STORAGE_BLOCK_LABEL.FontSize = "16"
  $STORAGE_BLOCK_LABEL.FontWeight = "Bold"

  # PAINEL
  $StackPanel = New-Object System.Windows.Controls.StackPanel

  $StackPanel.AddChild($NOTEBOOK_BLOCK)

  # ADD CPU AO PAINEL
  $StackPanel.AddChild($CPU_BLOCK_LABEL)
  $StackPanel.AddChild($CPU_BLOCK)
  $StackPanel.AddChild($CPU_CORES_BLOCK)
  $StackPanel.AddChild($CPU_SPEED_BLOCK)

  # ADD MEMÓRIA AO PAINEL
  $StackPanel.AddChild($MEMORY_BLOCK_LABEL)
  $StackPanel.AddChild($MEMORY_BLOCK)
  $StackPanel.AddChild($MEMORY_SLOTS_BLOCK)

  # TAMANHO CADA SLOT, SE HOUVER MAIS DE UM
  if ($MEMORY_SLOTS -gt 1) {
    $MEMORY_SLOTS_SIZE_BLOCK = New-Object System.Windows.Controls.TextBlock
    $MEMORY_SLOTS_SIZE_BLOCK.Text = $SLOT_SIZE_STRING
    $MEMORY_SLOTS_SIZE_BLOCK.margin = "30,0,15,0"
    $MEMORY_SLOTS_SIZE_BLOCK.FontSize = "16"

    $StackPanel.AddChild($MEMORY_SLOTS_SIZE_BLOCK)
  }

  $StackPanel.AddChild($MEMORY_SPEED_BLOCK)

  # ADD ARMAZENAMENTO AO PAINEL
  $StackPanel.AddChild($STORAGE_BLOCK_LABEL)

  # ARMAZENAMENTO CONTEÚDO
  foreach ($STORAGE_ITEM in $STORAGE_BASE) {
    $STORAGE_ITEM = $STORAGE_ITEM | Select-Object @{N="n";Expression={
      $(
        if (($_.PNPDeviceID -like "*NVME*") -or ($_.Model -like "*NVME*")) { "SSD NVMe "}
        elseif (($_.PNPDeviceID -like "*VEN_&*") -and ($_.Model -notlike "*NVME*")) {"SSD "}
        elseif(($_.PNPDeviceID -like "*VEN_SANDISK*") -or ($_.Model -like "*SanDisk SD*")) {"SSD "}
        elseif (($_.PNPDeviceID -like "*VEN_WDC*") -and ($_.PNPDeviceID -notlike "*NVME*")) {"HDD "}
        elseif($_.MediaType -like "*External*") {"HD Externo "}
        elseif($_.MediaType -like "*Removable*") {"Pen Drive "}
        elseif($_.PNPDeviceID -like "*RAID*") {"RAID "}
      ) + [string][math]::Round((($STORAGE_ITEM | Measure-Object -Property Size -sum).sum)/1GB, 2) + " GB"
      }
    } | Select-Object -ExpandProperty n

      $STORAGE_BLOCK = New-Object System.Windows.Controls.TextBlock
      $STORAGE_BLOCK.Text = $STORAGE_ITEM
      $STORAGE_BLOCK.margin = "30,0,15,0"
      $STORAGE_BLOCK.FontSize = "16"
      $StackPanel.AddChild($STORAGE_BLOCK)
  }

  $BUTTON_STACK_PANEL = New-Object System.Windows.Controls.StackPanel
  $BUTTON_STACK_PANEL.Orientation = "Horizontal"

  $BUTTON_AIM_CONTENT = New-Object System.Windows.Controls.TextBlock
  $BUTTON_AIM_CONTENT.Text = "  AIM  "
  $BUTTON_AIM_CONTENT.margin = "15,0,5,0"
  $BUTTON_AIM_CONTENT.FontSize = "18"
  $BUTTON_AIM_CONTENT.Foreground = "DarkRed"
  $BUTTON_AIM_CONTENT.HorizontalAlignment = "Center"
  $BUTTON_AIM_CONTENT.VerticalAlignment = "Center"
  # $BUTTON_AIM_CONTENT.FontWeight = "Bold"

  $BUTTON_AIM = New-Object System.Windows.Controls.Button
  $BUTTON_AIM.Content = $BUTTON_AIM_CONTENT
  $BUTTON_AIM.FontSize = 14
  # $BUTTON_AIM.HorizontalAlignment = "Left"
  $BUTTON_AIM.VerticalAlignment = "Center"
  $BUTTON_AIM.Height = 50
  $BUTTON_AIM.Width = 110
  $BUTTON_AIM.margin = "10,10,0,10"
  $BUTTON_AIM.Background = "Transparent"
  $BUTTON_AIM.BorderThickness = 2
  $BUTTON_AIM.Cursor = "Hand"
  # $BUTTON_AIM.Opacity = 0.9

  $BUTTON_AIM.Add_MouseEnter({
    # $BUTTON_AIM_CONTENT.FontSize = "20"
    $BUTTON_AIM_CONTENT.FontWeight = "Bold"
    # $BUTTON_AIM.Background = "Transparent"
    # $BUTTON_AIM_CONTENT.Foreground = "White"
  })

  $BUTTON_AIM.Add_MouseLeave({
    # $BUTTON_AIM.Background = "DarkRed"
    # $BUTTON_AIM_CONTENT.FontSize = "18"
    $BUTTON_AIM_CONTENT.FontWeight = "Normal"
    # $BUTTON_AIM_CONTENT.Foreground = "DarkRed"
  })


  $Params_Aim=@{
    Content="Nenhuma instalação do AIM encontrada."
    Title="Atenção"
    TitleBackground="Orange"
    TitleFontSize=22
    TitleFontWeight='Bold'
    TitleTextForeground='Black'
    ButtonType='None'
    ButtonTextForeground="Black"
    CustomButtons="Voltar"
    BorderThickness=2
    ShadowDepth=4
    # ContentBackground="Black"
  }

  $BUTTON_AIM.Add_Click({
    $AIM = Test-Path "C:\Program Files (x86)\Automatos\Desktop Agent\adacontrol.exe"
    if (!$AIM) {
      New-WPFMessageBox @Params_AIM
    } else {
      Start-Process -FilePath "C:\Program Files (x86)\Automatos\Desktop Agent\adacontrol.exe"
    }
  })

  $BUTTON_WINDOWS_CONTENT = New-Object System.Windows.Controls.TextBlock
  $BUTTON_WINDOWS_CONTENT.Text = "ATIVAÇÃO WINDOWS"
  $BUTTON_WINDOWS_CONTENT.margin = "15,0,15,0"
  $BUTTON_WINDOWS_CONTENT.FontSize = "18"
  $BUTTON_WINDOWS_CONTENT.Foreground = "DarkRed"
  $BUTTON_WINDOWS_CONTENT.HorizontalAlignment = "Center"
  $BUTTON_WINDOWS_CONTENT.VerticalAlignment = "Center"
  # $BUTTON_WINDOWS_CONTENT.FontWeight = "Bold"

  $BUTTON_WINDOWS = New-Object System.Windows.Controls.Button
  $BUTTON_WINDOWS.Content = $BUTTON_WINDOWS_CONTENT
  $BUTTON_WINDOWS.FontSize = 14
  # $BUTTON_WINDOWS.HorizontalAlignment = "Left"
  $BUTTON_WINDOWS.VerticalAlignment = "Center"
  $BUTTON_WINDOWS.Height = 50
  $BUTTON_WINDOWS.Width = 250
  $BUTTON_WINDOWS.margin = "0,10,0,10"
  $BUTTON_WINDOWS.Background = "Transparent"
  $BUTTON_WINDOWS.BorderThickness = 0
  $BUTTON_WINDOWS.Cursor = "Hand"
  # $BUTTON_WINDOWS.Opacity = 0.9

  $BUTTON_WINDOWS.Add_Click({
    Start-Process -FilePath "slui.exe"
  })

  $BUTTON_WINDOWS.Add_MouseEnter({
    # $BUTTON_WINDOWS_CONTENT.FontSize = "20"
    $BUTTON_WINDOWS_CONTENT.FontWeight = "Bold"
    # $BUTTON_WINDOWS.Background = "DarkRed"
    # $BUTTON_WINDOWS_CONTENT.Foreground = "White"
  })

  $BUTTON_WINDOWS.Add_MouseLeave({
    # $BUTTON_WINDOWS_CONTENT.FontSize = "18"
    $BUTTON_WINDOWS_CONTENT.FontWeight = "Normal"
    # $BUTTON_WINDOWS.Background = "Transparent"
    # $BUTTON_WINDOWS_CONTENT.Foreground = "DarkRed"
  })

  $BUTTON_DRIVERS_CONTENT = New-Object System.Windows.Controls.TextBlock
  $BUTTON_DRIVERS_CONTENT.Text = "DRIVERS"
  $BUTTON_DRIVERS_CONTENT.margin = "15,0,15,0"
  $BUTTON_DRIVERS_CONTENT.FontSize = "18"
  $BUTTON_DRIVERS_CONTENT.Foreground = "DarkRed"
  $BUTTON_DRIVERS_CONTENT.HorizontalAlignment = "Center"
  $BUTTON_DRIVERS_CONTENT.VerticalAlignment = "Center"
  # $BUTTON_DRIVERS_CONTENT.FontWeight = "Bold"

  $BUTTON_DRIVERS = New-Object System.Windows.Controls.Button
  $BUTTON_DRIVERS.Content = $BUTTON_DRIVERS_CONTENT
  $BUTTON_DRIVERS.FontSize = 14
  # $BUTTON_DRIVERS.HorizontalAlignment = "Left"
  $BUTTON_DRIVERS.VerticalAlignment = "Center"
  $BUTTON_DRIVERS.Height = 50
  $BUTTON_DRIVERS.Width = 110
  $BUTTON_DRIVERS.margin = "0,10,35,10"
  # $BUTTON_DRIVERS.padding = "0,0,50,0"
  $BUTTON_DRIVERS.Background = "Transparent"
  $BUTTON_DRIVERS.BorderThickness = 10
  $BUTTON_DRIVERS.Cursor = "Hand"
  $BUTTON_DRIVERS.BorderBrush = "Black"
  # $BUTTON_DRIVERS.Bord
  # $BUTTON_DRIVERS.Opacity = 1

  $BUTTON_DRIVERS.Add_MouseEnter({
    # $BUTTON_DRIVERS_CONTENT.FontSize = "20"
    # $BUTTON_DRIVERS.Background = "DarkRed"
    # $BUTTON_DRIVERS_CONTENT.Foreground = "White"
    $BUTTON_DRIVERS_CONTENT.FontWeight = "Bold"
  })

  $BUTTON_DRIVERS.Add_MouseLeave({
    # $BUTTON_DRIVERS_CONTENT.FontSize = "18"
    # $BUTTON_DRIVERS.Background = "Transparent"
    # $BUTTON_DRIVERS_CONTENT.Foreground = "DarkRed"
    $BUTTON_DRIVERS_CONTENT.FontWeight = "normal"
  })

  $BUTTON_DRIVERS.Add_Click({
    Start-Process -FilePath "devmgmt.msc"
  })

  # $BUTTON_STACK_PANEL.Background = "DarkRed"
  $BUTTON_STACK_PANEL.margin = "0,10,0,0"
  $BUTTON_STACK_PANEL.height = 50
  # # $BUTTON_STACK_PANEL.Opacity = 0.97
  # $BUTTON_STACK_PANEL.BorderThickness = 2
  # $DockPanel = New-object System.Windows.Controls.DockPanel
  # $DockPanel.LastChildFill = $False
  # $DockPanel.HorizontalAlignment = "Center"
  # $DockPanel.Width = "NaN"
  # $DockPanel.AddChild($BUTTON_AIM)
  # $DockPanel.AddChild($BUTTON_WINDOWS)
  # $DockPanel.AddChild($BUTTON_DRIVERS)
  # $DockPanel.Background = "Black"

  $BUTTON_STACK_PANEL.AddChild($BUTTON_AIM)
  $BUTTON_STACK_PANEL.AddChild($BUTTON_WINDOWS)
  $BUTTON_STACK_PANEL.AddChild($BUTTON_DRIVERS)


  # ADD IMAGEM AO PAINEL
  $StackPanel.Background="White"
  $StackPanel.AddChild($IMAGE)
  $StackPanel.AddChild($BUTTON_STACK_PANEL)

  # Parametros gerais
  $Params=@{
    Content=$StackPanel
    Title="VALIDAÇÃO"
    TitleBackground="DarkRed"
    TitleFontSize=22
    TitleFontWeight='Bold'
    TitleTextForeground='White'
    ButtonType='None'
    ButtonTextForeground="DarkRed"
    CustomButtons="OK"
    BorderThickness=2
    ShadowDepth=4
    ContentBackground="White"
  }

  New-WPFMessageBox @Params

  if ($WPFMessageBoxOutput -eq "OK") {
    exit
  }
}

# Resolução após clique nos botoões
