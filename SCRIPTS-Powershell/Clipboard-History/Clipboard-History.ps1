#Requires -Version 3.0

# Add-Type for GetClipboardSequenceNumber
Add-Type -Namespace Api -Name Win32 -MemberDefinition '
[DllImport("user32.dll")]
public static extern uint GetClipboardSequenceNumber();
'

$Runspacehash = [hashtable]::Synchronized(@{})
$Runspacehash.Host = $Host
$Runspacehash.runspace = [RunspaceFactory]::CreateRunspace()
$Runspacehash.runspace.ApartmentState = "STA"
$Runspacehash.runspace.Open()
$Runspacehash.runspace.SessionStateProxy.SetVariable("Runspacehash",$Runspacehash)
$Runspacehash.PowerShell = {Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase}.GetPowerShell()
$Runspacehash.PowerShell.Runspace = $Runspacehash.runspace
$Runspacehash.Handle = $Runspacehash.PowerShell.AddScript({
    Function Get-ClipBoard {
        [Windows.Clipboard]::GetText()
    }
    Function Set-ClipBoard {
        $CopiedText = $listbox.SelectedItems | Out-String
        [Windows.Clipboard]::SetText($CopiedText.Trim())
    }
    Function Clear-Viewer {
        [void]$Script:ObservableCollection.Clear()
        [Windows.Clipboard]::Clear()
    }
    Function Copy-T {
        [Windows.Clipboard]::SetText("t")
        if ($Script:ObservableCollection.Count -eq 0 -or "t" -ne $Script:ObservableCollection[0]) {
            [void]$Script:ObservableCollection.Insert(0, "t")
            if ($Script:ObservableCollection.Count -gt 20) {
                [void]$Script:ObservableCollection.RemoveAt(20)
            }
        }
    }

    #Build the GUI
    [xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Powershell Clipboard History Viewer" WindowStartupLocation = "CenterScreen" 
    Width = "350" Height = "425" ShowInTaskbar = "True" Background = "White">
    <Grid >
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.Resources>
            <Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
                <Setter Property="Background" Value="LightGray"/>
                <Setter Property="Foreground" Value="Black"/>
                <Style.Triggers>
                    <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                        <Setter Property="Background" Value="White"/>
                        <Setter Property="Foreground" Value="Black"/>
                    </Trigger>
                </Style.Triggers>
            </Style>
        </Grid.Resources>
        <Menu Width = 'Auto' HorizontalAlignment = 'Stretch' Grid.Row = '0'>
        <Menu.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> 
                <GradientStop Color='#C4CBD8' Offset='0' /> 
                <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> 
                <GradientStop Color='#C4CBD8' Offset='1' /> 
                </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Menu.Background>
            <MenuItem x:Name = 'FileMenu' Header = '_File'>
                <MenuItem x:Name = 'Clear_Menu' Header = '_Clear' />
            </MenuItem>
        </Menu>
        <Button x:Name="CopyTButton" Content="Copy 't' to Clipboard" Height="25" Grid.Row="1" Margin="5"/>
        <GroupBox Header = "Filter"  Grid.Row = '2' Background = "White">
            <TextBox x:Name="InputBox" Height = "25" Grid.Row="2" />
        </GroupBox>
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"  
        Grid.Row="4" Height = "Auto">
            <ListBox x:Name="listbox" AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}" 
            SelectionMode='Extended'>
            <ListBox.Template>
                <ControlTemplate TargetType="ListBox">
                    <Border BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderBrush}">
                        <ItemsPresenter/>
                    </Border>
                </ControlTemplate>
            </ListBox.Template>
            <ListBox.ContextMenu>
                <ContextMenu x:Name = 'ClipboardMenu'>
                    <MenuItem x:Name = 'Copy_Menu' Header = 'Copy'/>      
                    <MenuItem x:Name = 'Remove_Menu' Header = 'Remove'/>  
                </ContextMenu>
            </ListBox.ContextMenu>
            </ListBox>
        </ScrollViewer >
    </Grid>
</Window>
"@

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $Window=[Windows.Markup.XamlReader]::Load( $reader )

    #Connect to Controls
    $listbox = $Window.FindName('listbox')
    $InputBox = $Window.FindName('InputBox')
    $Clear_Menu = $Window.FindName('Clear_Menu')
    $Copy_Menu = $Window.FindName('Copy_Menu')
    $Remove_Menu = $Window.FindName('Remove_Menu')
    $CopyTButton = $Window.FindName('CopyTButton')

    #Observable Collection
    $Script:ObservableCollection = New-Object System.Collections.ObjectModel.ObservableCollection[String]

    #Collection View for filtering
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:ObservableCollection)
    $view.Filter = { param($item) $item -match $InputBox.Text }
    $listbox.ItemsSource = $view

    #Events
    $InputBox.Add_TextChanged({
        $view.Refresh()
    })

    $Copy_Menu.Add_Click({
        Set-ClipBoard
    })

    $Remove_Menu.Add_Click({
        $selected = @($listbox.SelectedItems)
        foreach ($item in $selected) {
            [void]$Script:ObservableCollection.Remove($item)
        }
    })

    $Clear_Menu.Add_Click({
        Clear-Viewer
    })

    $CopyTButton.Add_Click({
        Copy-T
    })

    $listbox.Add_MouseDoubleClick({
        if ($listbox.SelectedItem) {
            [Windows.Clipboard]::SetText($listbox.SelectedItem)
        }
    })

    #Timer for clipboard monitoring
    $Script:LastSeq = [Api.Win32]::GetClipboardSequenceNumber()
    $Timer = New-Object System.Windows.Threading.DispatcherTimer
    $Timer.Interval = [TimeSpan]::FromSeconds(1)
    $Timer.Add_Tick({
        $CurrentSeq = [Api.Win32]::GetClipboardSequenceNumber()
        if ($CurrentSeq -ne $Script:LastSeq) {
            $Script:LastSeq = $CurrentSeq
            $ClipboardText = Get-ClipBoard
            if ($ClipboardText -and ($Script:ObservableCollection.Count -eq 0 -or $ClipboardText -ne $Script:ObservableCollection[0])) {
                [void]$Script:ObservableCollection.Insert(0, $ClipboardText)
                if ($Script:ObservableCollection.Count -gt 20) {
                    [void]$Script:ObservableCollection.RemoveAt(20)
                }
            }
        }
    })
    $Timer.Start()

    #Initial clipboard check
    $InitialText = Get-ClipBoard
    if ($InitialText) {
        [void]$Script:ObservableCollection.Add($InitialText)
    }

    $window.ShowDialog() | Out-Null
}).BeginInvoke()
