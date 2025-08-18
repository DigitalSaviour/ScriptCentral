#Requires -Version 3.0

# Add required assemblies
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Drawing
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
$Runspacehash.PowerShell = {Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Drawing}.GetPowerShell()
$Runspacehash.PowerShell.Runspace = $Runspacehash.runspace
$Runspacehash.Handle = $Runspacehash.PowerShell.AddScript({
    # Helper functions
    Function Get-ClipBoard {
        try {
            if ([Windows.Clipboard]::ContainsText()) {
                return @{ Type = "Text"; Content = [Windows.Clipboard]::GetText() }
            }
            elseif ([Windows.Clipboard]::ContainsImage()) {
                $img = [Windows.Clipboard]::GetImage()
                if ($img) {
                    $tempPath = Join-Path $env:TEMP ("ClipboardImg_" + [Guid]::NewGuid() + ".png")
                    $bitmap = [System.Drawing.Bitmap]$img
                    $bitmap.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $bitmap.Dispose()
                    return @{ Type = "Image"; Content = $tempPath }
                }
            }
            elseif ([Windows.Clipboard]::ContainsFileDropList()) {
                $files = [Windows.Clipboard]::GetFileDropList() | Out-String
                return @{ Type = "Files"; Content = "[Files]`n$files" }
            }
        }
        catch {
            Write-Error "Error getting clipboard content: $_"
        }
        return $null
    }

    Function Set-ClipBoard {
        $selected = $listbox.SelectedItems
        if ($selected) {
            $item = $selected[0]
            try {
                if ($item.Type -eq "Text" -or $item.Type -eq "Files") {
                    [Windows.Clipboard]::SetText($item.Content)
                }
                elseif ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                    $img = [System.Drawing.Bitmap]::new($item.Content)
                    [Windows.Clipboard]::SetImage($img)
                    $img.Dispose()
                }
            }
            catch {
                Write-Error "Error setting clipboard: $_"
            }
        }
    }

    Function Clear-Viewer {
        try {
            foreach ($item in $Script:ObservableCollection) {
                if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                    Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                }
            }
            [void]$Script:ObservableCollection.Clear()
            [Windows.Clipboard]::Clear()
        }
        catch {
            Write-Error "Error clearing viewer: $_"
        }
    }

    Function Copy-T {
        try {
            [Windows.Clipboard]::SetText("t")
            if ($Script:ObservableCollection.Count -eq 0 -or $Script:ObservableCollection[0].Content -ne "t") {
                [void]$Script:ObservableCollection.Insert(0, [PSCustomObject]@{
                    Type = "Text"; Content = "t"; Timestamp = (Get-Date); IsPinned = $false
                })
                if ($Script:ObservableCollection.Count -gt 100) {
                    $item = $Script:ObservableCollection[$Script:ObservableCollection.Count - 1]
                    if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                        Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                    }
                    [void]$Script:ObservableCollection.RemoveAt($Script:ObservableCollection.Count - 1)
                }
            }
        }
        catch {
            Write-Error "Error copying 't': $_"
        }
    }

    Function Copy-S {
        try {
            [Windows.Clipboard]::SetText("s")
            if ($Script:ObservableCollection.Count -eq 0 -or $Script:ObservableCollection[0].Content -ne "s") {
                [void]$Script:ObservableCollection.Insert(0, [PSCustomObject]@{
                    Type = "Text"; Content = "s"; Timestamp = (Get-Date); IsPinned = $false
                })
                if ($Script:ObservableCollection.Count -gt 100) {
                    $item = $Script:ObservableCollection[$Script:ObservableCollection.Count - 1]
                    if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                        Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                    }
                    [void]$Script:ObservableCollection.RemoveAt($Script:ObservableCollection.Count - 1)
                }
            }
        }
        catch {
            Write-Error "Error copying 's': $_"
        }
    }

    Function Pin-Item {
        try {
            $selected = $listbox.SelectedItems
            if ($selected) {
                foreach ($item in $selected) {
                    $item.IsPinned = -not $item.IsPinned
                }
                $view.SortDescriptions.Clear()
                $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("IsPinned", "Descending")))
                $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("Timestamp", "Descending")))
                $view.Refresh()
            }
        }
        catch {
            Write-Error "Error pinning item: $_"
        }
    }

    Function Save-History {
        try {
            $items = $Script:ObservableCollection | Where-Object { $_.Type -eq "Text" -or $_.Type -eq "Files" } | Select-Object Type, Content, Timestamp, IsPinned
            $items | ConvertTo-Json | Out-File "$env:APPDATA\ClipboardHistory.json" -Force
        }
        catch {
            Write-Error "Error saving history: $_"
        }
    }

    Function Load-History {
        try {
            $historyFile = "$env:APPDATA\ClipboardHistory.json"
            if (Test-Path $historyFile) {
                $items = Get-Content $historyFile -Raw | ConvertFrom-Json
                foreach ($item in $items) {
                    [void]$Script:ObservableCollection.Add([PSCustomObject]@{
                        Type = $item.Type
                        Content = $item.Content
                        Timestamp = [DateTime]$item.Timestamp
                        IsPinned = [bool]$item.IsPinned
                    })
                }
            }
        }
        catch {
            Write-Error "Error loading history: $_"
        }
    }

    #Build the GUI
    [xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Clipboard History Manager" WindowStartupLocation="CenterScreen" 
    Width="400" Height="500" ShowInTaskbar="True" ResizeMode="CanResize" MinWidth="300" MinHeight="400"
    Background="#F5F7FA">
    <Window.Resources>
        <SolidColorBrush x:Key="PrimaryBackground" Color="#F5F7FA"/>
        <SolidColorBrush x:Key="SecondaryBackground" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="ButtonBackground" Color="#0078D4"/>
        <SolidColorBrush x:Key="ButtonHoverBackground" Color="#005EA6"/>
        <SolidColorBrush x:Key="ButtonPressedBackground" Color="#004B8D"/>
        <SolidColorBrush x:Key="TextForeground" Color="#000000"/>
        <SolidColorBrush x:Key="SecondaryTextForeground" Color="#666666"/>
        <SolidColorBrush x:Key="HighlightBackground" Color="#E8F0FE"/>
        <SolidColorBrush x:Key="BorderColor" Color="#E0E0E0"/>
        <Style x:Key="ButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ButtonBackground}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{DynamicResource ButtonHoverBackground}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{DynamicResource ButtonPressedBackground}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ListBoxItemStyle" TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="Margin" Value="2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{DynamicResource BorderColor}"
                                BorderThickness="0,0,0,1"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="{DynamicResource HighlightBackground}"/>
                                <Setter Property="Foreground" Value="{DynamicResource ButtonBackground}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{DynamicResource HighlightBackground}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <DataTemplate x:Key="ListBoxItemTemplate">
            <StackPanel Orientation="Horizontal">
                <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE734;" Margin="0,0,5,0">
                    <TextBlock.Style>
                        <Style TargetType="TextBlock">
                            <Setter Property="Visibility" Value="Collapsed"/>
                            <Setter Property="Foreground" Value="{DynamicResource TextForeground}"/>
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding IsPinned}" Value="True">
                                    <Setter Property="Visibility" Value="Visible"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </TextBlock.Style>
                </TextBlock>
                <StackPanel>
                    <TextBlock FontFamily="Segoe UI" FontSize="13" Foreground="{DynamicResource TextForeground}">
                        <Run Text="{Binding Content}"/>
                    </TextBlock>
                    <TextBlock Text="{Binding Timestamp, StringFormat='Copied: {0:yyyy-MM-dd HH:mm:ss}'}" 
                               FontFamily="Segoe UI" FontSize="10" Foreground="{DynamicResource SecondaryTextForeground}"/>
                </StackPanel>
                <Image Source="{Binding Content}" MaxHeight="50" Margin="5,0,0,0">
                    <Image.Style>
                        <Style TargetType="Image">
                            <Setter Property="Visibility" Value="Collapsed"/>
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding Type}" Value="Image">
                                    <Setter Property="Visibility" Value="Visible"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </Image.Style>
                </Image>
            </StackPanel>
        </DataTemplate>
        <BooleanToVisibilityConverter x:Key="BooleanToVisibilityConverter"/>
    </Window.Resources>
    <Border Background="{DynamicResource SecondaryBackground}" CornerRadius="8" Margin="10" BorderBrush="{DynamicResource BorderColor}" BorderThickness="1">
        <Border.Effect>
            <DropShadowEffect ShadowDepth="2" Direction="315" Color="#A0A0A0" Opacity="0.3" BlurRadius="8"/>
        </Border.Effect>
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Menu Width="Auto" HorizontalAlignment="Stretch" Grid.Row="0" Background="Transparent" Foreground="{DynamicResource TextForeground}">
                <MenuItem x:Name="FileMenu" Header="_File" FontFamily="Segoe UI" FontSize="14" Foreground="{DynamicResource TextForeground}">
                    <MenuItem x:Name="Clear_Menu" Header="_Clear" FontFamily="Segoe UI" Foreground="{DynamicResource TextForeground}">
                        <MenuItem.Icon>
                            <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE74D;" Foreground="{DynamicResource TextForeground}"/>
                        </MenuItem.Icon>
                    </MenuItem>
                    <MenuItem x:Name="Theme_Menu" Header="_Toggle Theme" FontFamily="Segoe UI" Foreground="{DynamicResource TextForeground}">
                        <MenuItem.Icon>
                            <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE771;" Foreground="{DynamicResource TextForeground}"/>
                        </MenuItem.Icon>
                    </MenuItem>
                </MenuItem>
            </Menu>
            <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,5" HorizontalAlignment="Center">
                <Button x:Name="CopyTButton" Style="{StaticResource ButtonStyle}">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE8C8;" Margin="0,0,5,0" Foreground="White"/>
                        <TextBlock Text="Copy 't' to Clipboard" Foreground="White"/>
                    </StackPanel>
                </Button>
                <Button x:Name="CopySButton" Style="{StaticResource ButtonStyle}">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE8C8;" Margin="0,0,5,0" Foreground="White"/>
                        <TextBlock Text="Copy 's' to Clipboard" Foreground="White"/>
                    </StackPanel>
                </Button>
            </StackPanel>
            <GroupBox Header="Filter" Grid.Row="2" Background="{DynamicResource SecondaryBackground}" Margin="0,5" FontFamily="Segoe UI" Foreground="{DynamicResource TextForeground}">
                <TextBox x:Name="InputBox" Height="25" FontFamily="Segoe UI" FontSize="13" Margin="5" Foreground="{DynamicResource TextForeground}" Background="{DynamicResource SecondaryBackground}"/>
            </GroupBox>
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Grid.Row="4" Margin="0,5">
                <ListBox x:Name="listbox" SelectionMode="Extended" ItemTemplate="{StaticResource ListBoxItemTemplate}" ItemContainerStyle="{StaticResource ListBoxItemStyle}" Foreground="{DynamicResource TextForeground}" Background="{DynamicResource SecondaryBackground}">
                    <ListBox.Template>
                        <ControlTemplate TargetType="ListBox">
                            <Border Background="{TemplateBinding Background}" 
                                    BorderBrush="{DynamicResource BorderColor}" 
                                    BorderThickness="1" 
                                    CornerRadius="4">
                                <ItemsPresenter/>
                            </Border>
                        </ControlTemplate>
                    </ListBox.Template>
                    <ListBox.ContextMenu>
                        <ContextMenu x:Name="ClipboardMenu" FontFamily="Segoe UI" Foreground="{DynamicResource TextForeground}" Background="{DynamicResource SecondaryBackground}">
                            <MenuItem x:Name="Copy_Menu" Header="Copy" Foreground="{DynamicResource TextForeground}">
                                <MenuItem.Icon>
                                    <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE8C8;" Foreground="{DynamicResource TextForeground}"/>
                                </MenuItem.Icon>
                            </MenuItem>
                            <MenuItem x:Name="Remove_Menu" Header="Remove" Foreground="{DynamicResource TextForeground}">
                                <MenuItem.Icon>
                                    <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE74D;" Foreground="{DynamicResource TextForeground}"/>
                                </MenuItem.Icon>
                            </MenuItem>
                            <MenuItem x:Name="Pin_Menu" Header="Pin/Unpin" Foreground="{DynamicResource TextForeground}">
                                <MenuItem.Icon>
                                    <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE718;" Foreground="{DynamicResource TextForeground}"/>
                                </MenuItem.Icon>
                            </MenuItem>
                        </ContextMenu>
                    </ListBox.ContextMenu>
                </ListBox>
            </ScrollViewer>
        </Grid>
    </Border>
</Window>
"@

    try {
        # Load XAML
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $Window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "XAML loaded successfully."

        # Connect to Controls
        $listbox = $Window.FindName('listbox')
        $InputBox = $Window.FindName('InputBox')
        $Clear_Menu = $Window.FindName('Clear_Menu')
        $Copy_Menu = $Window.FindName('Copy_Menu')
        $Remove_Menu = $Window.FindName('Remove_Menu')
        $Pin_Menu = $Window.FindName('Pin_Menu')
        $Theme_Menu = $Window.FindName('Theme_Menu')
        $CopyTButton = $Window.FindName('CopyTButton')
        $CopySButton = $Window.FindName('CopySButton')
        Write-Host "Controls connected successfully."

        # Observable Collection
        $Script:ObservableCollection = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]

        # Collection View for filtering
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:ObservableCollection)
        $view.Filter = { 
            param($item) 
            try {
                if ($InputBox.Text) {
                    $item.Content -imatch [regex]::Escape($InputBox.Text)
                } else {
                    $true
                }
            }
            catch {
                Write-Error "Error in filter: $_"
                $true
            }
        }
        $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("IsPinned", "Descending")))
        $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("Timestamp", "Descending")))
        $listbox.ItemsSource = $view
        Write-Host "Collection view set up successfully."

        # Theme switching
        $Script:IsDarkTheme = $false
        Function Toggle-Theme {
            try {
                $Script:IsDarkTheme = -not $Script:IsDarkTheme
                $resources = $Window.Resources
                if ($Script:IsDarkTheme) {
                    $resources["PrimaryBackground"].Color = "#2D2D2D"
                    $resources["SecondaryBackground"].Color = "#1F1F1F"
                    $resources["TextForeground"].Color = "#FFFFFF"
                    $resources["SecondaryTextForeground"].Color = "#A0A0A0"
                    $resources["HighlightBackground"].Color = "#3A3A3A"
                    $resources["BorderColor"].Color = "#4A4A4A"
                } else {
                    $resources["PrimaryBackground"].Color = "#F5F7FA"
                    $resources["SecondaryBackground"].Color = "#FFFFFF"
                    $resources["TextForeground"].Color = "#000000"
                    $resources["SecondaryTextForeground"].Color = "#666666"
                    $resources["HighlightBackground"].Color = "#E8F0FE"
                    $resources["BorderColor"].Color = "#E0E0E0"
                }
                Write-Host "Theme toggled to $($Script:IsDarkTheme ? 'Dark' : 'Light')."
            }
            catch {
                Write-Error "Error toggling theme: $_"
            }
        }

        # Keyboard shortcuts
        $Window.Add_KeyDown({
            param($sender, $e)
            try {
                if ($e.Key -eq "T" -and $e.KeyboardDevice.Modifiers -eq "Ctrl") { Copy-T }
                elseif ($e.Key -eq "S" -and $e.KeyboardDevice.Modifiers -eq "Ctrl") { Copy-S }
                elseif ($e.Key -eq "C" -and $e.KeyboardDevice.Modifiers -eq "Ctrl") { Set-ClipBoard }
                elseif ($e.Key -eq "D" -and $e.KeyboardDevice.Modifiers -eq "Ctrl") {
                    $selected = @($listbox.SelectedItems)
                    foreach ($item in $selected) {
                        if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                            Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                        }
                        [void]$Script:ObservableCollection.Remove($item)
                    }
                }
                elseif ($e.Key -eq "P" -and $e.KeyboardDevice.Modifiers -eq "Ctrl") { Pin-Item }
            }
            catch {
                Write-Error "Error handling key press: $_"
            }
        })

        # Events
        $InputBox.Add_TextChanged({ try { $view.Refresh() } catch { Write-Error "Error refreshing view: $_" } })
        $Copy_Menu.Add_Click({ Set-ClipBoard })
        $Remove_Menu.Add_Click({
            try {
                $selected = @($listbox.SelectedItems)
                foreach ($item in $selected) {
                    if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                        Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                    }
                    [void]$Script:ObservableCollection.Remove($item)
                }
            } catch { Write-Error "Error removing item: $_" }
        })
        $Pin_Menu.Add_Click({ Pin-Item })
        $Clear_Menu.Add_Click({ Clear-Viewer })
        $Theme_Menu.Add_Click({ Toggle-Theme })
        $CopyTButton.Add_Click({ Copy-T })
        $CopySButton.Add_Click({ Copy-S })
        $listbox.Add_MouseDoubleClick({ if ($listbox.SelectedItem) { Set-ClipBoard } })
        $Window.Add_Closing({
            Save-History
            foreach ($item in $Script:ObservableCollection) {
                if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                    Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                }
            }
        })

        # Timer for clipboard monitoring
        $Script:LastSeq = [Api.Win32]::GetClipboardSequenceNumber()
        $Timer = New-Object System.Windows.Threading.DispatcherTimer
        $Timer.Interval = [TimeSpan]::FromSeconds(1)
        $Timer.Add_Tick({
            try {
                $CurrentSeq = [Api.Win32]::GetClipboardSequenceNumber()
                if ($CurrentSeq -ne $Script:LastSeq) {
                    $Script:LastSeq = $CurrentSeq
                    $ClipboardItem = Get-ClipBoard
                    if ($ClipboardItem -and ($Script:ObservableCollection.Count -eq 0 -or $ClipboardItem.Content -ne $Script:ObservableCollection[0].Content)) {
                        [void]$Script:ObservableCollection.Insert(0, [PSCustomObject]@{
                            Type = $ClipboardItem.Type
                            Content = $ClipboardItem.Content
                            Timestamp = (Get-Date)
                            IsPinned = $false
                        })
                        if ($Script:ObservableCollection.Count -gt 100) {
                            $item = $Script:ObservableCollection[$Script:ObservableCollection.Count - 1]
                            if ($item.Type -eq "Image" -and (Test-Path $item.Content)) {
                                Remove-Item $item.Content -Force -ErrorAction SilentlyContinue
                            }
                            [void]$Script:ObservableCollection.RemoveAt($Script:ObservableCollection.Count - 1)
                        }
                    }
                }
            }
            catch {
                Write-Error "Error in clipboard monitoring: $_"
            }
        })
        $Timer.Start()

        # Load history and initial clipboard
        try {
            Load-History
            $InitialItem = Get-ClipBoard
            if ($InitialItem) {
                [void]$Script:ObservableCollection.Add([PSCustomObject]@{
                    Type = $InitialItem.Type
                    Content = $InitialItem.Content
                    Timestamp = (Get-Date)
                    IsPinned = $false
                })
            }
        }
        catch {
            Write-Error "Error loading initial data: $_"
        }

        # Show the window
        try {
            $Window.ShowDialog() | Out-Null
            Write-Host "Window displayed successfully."
        }
        catch {
            Write-Error "Failed to show window: $_"
        }
    }
    catch {
        Write-Error "Fatal error during initialization: $_"
    }
}).BeginInvoke()
