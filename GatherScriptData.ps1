<#
.SYNOPSIS
    Retrieves and displays script execution reports from the NinjaOne API, with filtering, false success detection, and performance optimizations for large-scale environments.

.DESCRIPTION
    This script fetches script execution activities from the NinjaOne platform for environments with up to 70,000 devices. It provides a WPF-based user interface to filter by Device Name, Script Name, Result, and Date Range, and displays results in a DataGrid. The script identifies "SUCCESS" entries with errors in the output (e.g., "Cannot find any service with service name 'ADSync'") by searching for error patterns, highlighting them in orange with a tooltip, while maintaining green for true successes and red for failures. Performance optimizations include device pre-filtering, device caching, no API call delays, lazy timestamp conversion, and a default 1hr date range. Features include CSV export and a device cache clearing option.

    Key Features:
    - Filters: Device Name, Script Name, Result (ALL, SUCCESS, FAILURE), Date Range (1hr, 12hr, 1day, 7day, 30day, default 1hr).
    - False success detection: Flags "SUCCESS" entries with errors in output, displays in orange with tooltip.
    - Device cache at C:\NinjaCache\devices_cache.json to avoid repeated API calls.
    - Performance optimizations: Device pre-filtering, no API call delays, lazy timestamp conversion.
    - Export to CSV and device cache clearing button.

.EXAMPLE
    .\ScriptExecutionsReport.ps1
    Launches the script, allowing users to connect to the NinjaOne API, apply filters, generate reports, and export results to CSV.

.NOTES
    - Requires PowerShell 5.1 or later with WPF support.
    - Ensure write permissions to C:\NinjaCache for device caching.
    - Error patterns for false success detection can be customized in the script.

     **Disclaimer**: This script is provided "as is" under the MIT License. 
    Use at your own risk; the author is not responsible for any damages or issues arising from its use. 
    Licensed under MIT: Copyright (c) Spencer Heath. Permission is granted to use, copy, modify, and distribute this software freely, 
    provided the original copyright and this notice are retained. See https://opensource.org/licenses/MIT for full details.
#>
# Requires -Version 5.1
Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
Add-Type -AssemblyName System.IO

# Cache file paths (only for devices)
$cacheDir = "C:\NinjaCache"
$deviceCacheFile = Join-Path $cacheDir "devices_cache.json"

# Ensure the cache directory exists
if (-not (Test-Path $cacheDir)) {
    try {
        New-Item -ItemType Directory -Path $cacheDir -ErrorAction Stop | Out-Null
    } catch {
        Write-Error "Failed to create cache directory at $cacheDir. Error: $_"
        exit 1
    }
}

# Define error patterns for detecting false successes
$errorPatterns = @(
    "Cannot find",
    "Error:",
    "Exception:",
    "At C:\\",
    "CategoryInfo",
    "FullyQualifiedErrorId",
    "Failed to"
)

# Define WPF XAML with Updated Result Column
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="NinjaOne Script Executions Report" Height="780" Width="1550" WindowStartupLocation="CenterScreen" Background="#1E1E1E">
    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="300"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar for API Settings -->
        <Border Grid.Column="0" Background="#2D2D2D" Margin="0,0,15,0" BorderBrush="#444444" BorderThickness="1" Padding="10">
            <StackPanel>
                <Label Content="API Settings" Foreground="White" FontSize="16" Margin="0,0,0,10"/>
                <Label Content="NinjaOne Instance:" Foreground="White" FontSize="12"/>
                <ComboBox x:Name="InstanceComboBox" Width="250" Height="30" Margin="0,0,0,10" SelectedIndex="0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555">
                    <ComboBox.Resources>
                        <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#3C3C3C"/>
                        <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#3C3C3C"/>
                        <Style TargetType="ComboBox">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="BorderBrush" Value="#555555"/>
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="ComboBox">
                                        <Grid>
                                            <ToggleButton x:Name="ToggleButton" Background="#3C3C3C" BorderBrush="#555555" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                                <ToggleButton.ContentTemplate>
                                                    <DataTemplate>
                                                        <TextBlock Text="{Binding}" Foreground="White" Margin="3"/>
                                                    </DataTemplate>
                                                </ToggleButton.ContentTemplate>
                                            </ToggleButton>
                                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="3,3,23,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                                <Border x:Name="DropDownBorder" Background="#3C3C3C" BorderBrush="#555555" BorderThickness="1">
                                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                                    </ScrollViewer>
                                                </Border>
                                            </Popup>
                                            <Canvas x:Name="PART_DropDownGlyph" Width="10" Height="5" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,5,0">
                                                <Path x:Name="Path" Width="10" Height="5" Data="M 0 0 L 5 5 L 10 0 Z" Fill="White"/>
                                            </Canvas>
                                        </Grid>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsDropDownOpen" Value="True">
                                                <Setter TargetName="Path" Property="Data" Value="M 0 5 L 5 0 L 10 5 Z"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                        <Style TargetType="ComboBoxItem">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="BorderBrush" Value="#555555"/>
                            <Style.Triggers>
                                <Trigger Property="IsHighlighted" Value="True">
                                    <Setter Property="Background" Value="#555555"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </ComboBox.Resources>
                    <ComboBoxItem Content="app.ninjarmm.com"/>
                    <ComboBoxItem Content="us2.ninjarmm.com"/>
                    <ComboBoxItem Content="eu.ninjarmm.com"/>
                    <ComboBoxItem Content="ca.ninjarmm.com"/>
                    <ComboBoxItem Content="oc.ninjarmm.com"/>
                </ComboBox>
                <Label Content="Client ID:" Foreground="White" FontSize="12"/>
                <TextBox x:Name="ClientIdTextBox" Width="250" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Label Content="Client Secret:" Foreground="White" FontSize="12"/>
                <TextBox x:Name="ClientSecretTextBox" Width="250" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Button x:Name="ConnectButton" Content="Connect to API" Width="120" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Button x:Name="ClearCacheButton" Content="Clear Device Cache" Width="120" Height="30" Margin="0,0,0,10" Background="#FF4444" Foreground="White" BorderBrush="#555555"/>
                <Ellipse x:Name="ConnectionIndicator" Width="15" Height="15" Fill="Red" HorizontalAlignment="Left"/>
            </StackPanel>
        </Border>

        <!-- Main Content -->
        <Grid Grid.Column="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Label Grid.Row="0" Content="Script Executions Report" Foreground="White" FontSize="20" FontWeight="Bold" Margin="0,0,0,15"/>

            <!-- Filters -->
            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
                <Label Content="Device Name:" Foreground="White" FontSize="12" VerticalAlignment="Center"/>
                <TextBox x:Name="DeviceNameFilter" Width="200" Height="30" Margin="5,0,15,0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Label Content="Script Name:" Foreground="White" FontSize="12" VerticalAlignment="Center"/>
                <TextBox x:Name="ScriptNameFilter" Width="200" Height="30" Margin="5,0,15,0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Label Content="Result:" Foreground="White" FontSize="12" VerticalAlignment="Center"/>
                <ComboBox x:Name="ResultFilter" Width="150" Height="30" Margin="5,0,15,0" SelectedIndex="0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555">
                    <ComboBox.Resources>
                        <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#3C3C3C"/>
                        <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#3C3C3C"/>
                        <Style TargetType="ComboBox">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="BorderBrush" Value="#555555"/>
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="ComboBox">
                                        <Grid>
                                            <ToggleButton x:Name="ToggleButton" Background="#3C3C3C" BorderBrush="#555555" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                                <ToggleButton.ContentTemplate>
                                                    <DataTemplate>
                                                        <TextBlock Text="{Binding}" Foreground="White" Margin="3"/>
                                                    </DataTemplate>
                                                </ToggleButton.ContentTemplate>
                                            </ToggleButton>
                                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="3,3,23,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                                <Border x:Name="DropDownBorder" Background="#3C3C3C" BorderBrush="#555555" BorderThickness="1">
                                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                                    </ScrollViewer>
                                                </Border>
                                            </Popup>
                                            <Canvas x:Name="PART_DropDownGlyph" Width="10" Height="5" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,5,0">
                                                <Path x:Name="Path" Width="10" Height="5" Data="M 0 0 L 5 5 L 10 0 Z" Fill="White"/>
                                            </Canvas>
                                        </Grid>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsDropDownOpen" Value="True">
                                                <Setter TargetName="Path" Property="Data" Value="M 0 5 L 5 0 L 10 5 Z"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                        <Style TargetType="ComboBoxItem">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="BorderBrush" Value="#555555"/>
                            <Style.Triggers>
                                <Trigger Property="IsHighlighted" Value="True">
                                    <Setter Property="Background" Value="#555555"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </ComboBox.Resources>
                    <ComboBoxItem Content="ALL"/>
                    <ComboBoxItem Content="SUCCESS"/>
                    <ComboBoxItem Content="FAILURE"/>
                </ComboBox>
                <Label Content="Date Range:" Foreground="White" FontSize="12" VerticalAlignment="Center"/>
                <ComboBox x:Name="DateRangeFilter" Width="150" Height="30" Margin="5,0,15,0" SelectedIndex="0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555">
                    <ComboBox.Resources>
                        <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#3C3C3C"/>
                        <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#3C3C3C"/>
                        <Style TargetType="ComboBox">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="BorderBrush" Value="#555555"/>
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="ComboBox">
                                        <Grid>
                                            <ToggleButton x:Name="ToggleButton" Background="#3C3C3C" BorderBrush="#555555" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                                <ToggleButton.ContentTemplate>
                                                    <DataTemplate>
                                                        <TextBlock Text="{Binding}" Foreground="White" Margin="3"/>
                                                    </DataTemplate>
                                                </ToggleButton.ContentTemplate>
                                            </ToggleButton>
                                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="3,3,23,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                                <Border x:Name="DropDownBorder" Background="#3C3C3C" BorderBrush="#555555" BorderThickness="1">
                                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                                    </ScrollViewer>
                                                </Border>
                                            </Popup>
                                            <Canvas x:Name="PART_DropDownGlyph" Width="10" Height="5" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,5,0">
                                                <Path x:Name="Path" Width="10" Height="5" Data="M 0 0 L 5 5 L 10 0 Z" Fill="White"/>
                                            </Canvas>
                                        </Grid>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsDropDownOpen" Value="True">
                                                <Setter TargetName="Path" Property="Data" Value="M 0 5 L 5 0 L 10 5 Z"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                        <Style TargetType="ComboBoxItem">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="BorderBrush" Value="#555555"/>
                            <Style.Triggers>
                                <Trigger Property="IsHighlighted" Value="True">
                                    <Setter Property="Background" Value="#555555"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </ComboBox.Resources>
                    <ComboBoxItem Content="1hr"/>
                    <ComboBoxItem Content="12hr"/>
                    <ComboBoxItem Content="1day"/>
                    <ComboBoxItem Content="7day"/>
                    <ComboBoxItem Content="30day"/>
                </ComboBox>
                <CheckBox x:Name="ShowColorsCheckBox" Content="Show Result Colors" Foreground="White" FontSize="12" VerticalAlignment="Center" Margin="5,0,0,0" IsChecked="True"/>
            </StackPanel>

            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,15">
                <Button x:Name="GenerateButton" Content="Generate Report" Width="120" Height="30" Margin="15,0,0,0" Background="#0078D4" Foreground="White" BorderBrush="#555555"/>
                <Button x:Name="ExportButton" Content="Export to CSV" Width="120" Height="30" Margin="15,0,0,0" Background="#28A745" Foreground="White" BorderBrush="#555555" IsEnabled="False"/>
            </StackPanel>

            <DataGrid x:Name="ReportDataGrid" Grid.Row="3" Margin="0,0,0,15" AutoGenerateColumns="False" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" HeadersVisibility="Column" GridLinesVisibility="None" VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling" EnableRowVirtualization="True">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="*" />
                    <DataGridTextColumn Header="Script Name" Binding="{Binding ScriptName}" Width="*" />
                    <DataGridTextColumn Header="Result" Binding="{Binding Result}" Width="*">
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="ToolTip" Value="{Binding ErrorMessage}"/>
                                <Style.Triggers>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Result}" Value="SUCCESS"/>
                                            <Condition Binding="{Binding ErrorFlag}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Foreground" Value="Orange"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Result}" Value="SUCCESS"/>
                                            <Condition Binding="{Binding ErrorFlag}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Foreground" Value="Green"/>
                                    </MultiDataTrigger>
                                    <DataTrigger Binding="{Binding Result}" Value="FAILURE">
                                        <Setter Property="Foreground" Value="Red"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding ElementName=ShowColorsCheckBox, Path=IsChecked}" Value="False">
                                        <Setter Property="Foreground" Value="White"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Output" Binding="{Binding Output}" Width="*" />
                    <DataGridTextColumn Header="Timestamp" Binding="{Binding Timestamp}" Width="*" />
                </DataGrid.Columns>
                <DataGrid.Resources>
                    <Style TargetType="DataGrid">
                        <Setter Property="Background" Value="#3C3C3C"/>
                        <Setter Property="RowBackground" Value="#3C3C3C"/>
                        <Setter Property="AlternatingRowBackground" Value="#3C3C3C"/>
                        <Setter Property="BorderBrush" Value="#555555"/>
                    </Style>
                    <Style TargetType="DataGridColumnHeader">
                        <Setter Property="Background" Value="#0078D4"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="FontWeight" Value="Bold"/>
                        <Setter Property="Padding" Value="5"/>
                        <Setter Property="BorderBrush" Value="#555555"/>
                        <Setter Property="BorderThickness" Value="0,0,1,1"/>
                    </Style>
                    <Style TargetType="DataGridRow">
                        <Setter Property="Background" Value="#3C3C3C"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Style.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#555555"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#555555"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                    <Style TargetType="DataGridCell">
                        <Setter Property="Background" Value="#3C3C3C"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="BorderThickness" Value="0"/>
                        <Setter Property="Padding" Value="5"/>
                        <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#555555"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </DataGrid.Resources>
            </DataGrid>

            <Label Grid.Row="4" Content="Log:" Foreground="White" FontSize="14" Margin="0,0,0,5"/>
            <TextBox x:Name="LogTextBox" Grid.Row="5" Height="150" AcceptsReturn="True" VerticalScrollBarVisibility="Visible" IsReadOnly="True" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
        </Grid>
    </Grid>
</Window>
"@

# Create WPF Window
try {
    $window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
} catch {
    Write-Error "Failed to load XAML: $_"
    exit 1
}

# Get UI Elements
$instanceComboBox = $window.FindName("InstanceComboBox")
$clientIdTextBox = $window.FindName("ClientIdTextBox")
$clientSecretTextBox = $window.FindName("ClientSecretTextBox")
$connectButton = $window.FindName("ConnectButton")
$connectionIndicator = $window.FindName("ConnectionIndicator")
$deviceNameFilter = $window.FindName("DeviceNameFilter")
$scriptNameFilter = $window.FindName("ScriptNameFilter")
$resultFilter = $window.FindName("ResultFilter")
$dateRangeFilter = $window.FindName("DateRangeFilter")
$showColorsCheckBox = $window.FindName("ShowColorsCheckBox")
$clearCacheButton = $window.FindName("ClearCacheButton")
$generateButton = $window.FindName("GenerateButton")
$exportButton = $window.FindName("ExportButton")
$reportDataGrid = $window.FindName("ReportDataGrid")
$logTextBox = $window.FindName("LogTextBox")

# Check if all elements were found
if (-not $window -or -not $instanceComboBox -or -not $clientIdTextBox -or -not $clientSecretTextBox -or -not $connectButton -or -not $connectionIndicator -or -not $deviceNameFilter -or -not $scriptNameFilter -or -not $resultFilter -or -not $dateRangeFilter -or -not $showColorsCheckBox -or -not $clearCacheButton -or -not $generateButton -or -not $exportButton -or -not $reportDataGrid -or -not $logTextBox) {
    Write-Error "One or more UI elements could not be found."
    exit 1
}

# Global Variables
$global:DeviceLookup = @{}
$global:AccessToken = $null

# Function to Load Devices from Cache
function Load-DeviceCache {
    if (Test-Path $deviceCacheFile) {
        $cache = Get-Content $deviceCacheFile | ConvertFrom-Json
        $cacheAge = [datetime]::Parse($cache.Timestamp)
        if (((Get-Date) - $cacheAge).TotalHours -gt 24) {
            $logTextBox.Text += "Device cache is older than 24 hours, refreshing...`r`n"
            return $null
        }
        $global:DeviceLookup = @{}
        foreach ($device in $cache.Devices) {
            $global:DeviceLookup[$device.id] = $device.systemName
        }
        $logTextBox.Text += "Loaded $($global:DeviceLookup.Count) devices from cache at $deviceCacheFile`r`n"
        return $cache.Timestamp
    }
    return $null
}

# Function to Save Devices to Cache
function Save-DeviceCache {
    $cache = @{
        Timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        Devices = @($global:DeviceLookup.GetEnumerator() | ForEach-Object { @{ id = $_.Key; systemName = $_.Value } })
    }
    try {
        $cache | ConvertTo-Json -Depth 10 -ErrorAction Stop | Set-Content $deviceCacheFile -ErrorAction Stop
        $logTextBox.Text += "Saved $($global:DeviceLookup.Count) devices to cache at $deviceCacheFile`r`n"
    } catch {
        $logTextBox.Text += "Failed to save device cache at $deviceCacheFile. Error: $_`r`n"
    }
}

# Function to Detect Errors in Output
function Detect-ErrorInOutput($output) {
    foreach ($pattern in $errorPatterns) {
        if ($output -match [regex]::Escape($pattern)) {
            # Extract the line containing the error pattern
            $lines = $output -split "`r`n"
            $errorLine = $lines | Where-Object { $_ -match [regex]::Escape($pattern) } | Select-Object -First 1
            return $true, $errorLine
        }
    }
    return $false, $null
}

# Function to Authenticate
function Connect-NinjaOneAPI {
    $NinjaOneInstance = $instanceComboBox.SelectedItem.Content
    $NinjaOneClientId = $clientIdTextBox.Text
    $NinjaOneClientSecret = $clientSecretTextBox.Text

    $authHeaders = @{
        "accept" = "application/json"
        "Content-Type" = "application/x-www-form-urlencoded"
    }
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $NinjaOneClientId
        client_secret = $NinjaOneClientSecret
        scope         = "monitoring management"
    }

    try {
        $authResponse = Invoke-RestMethod -Uri "https://$NinjaOneInstance/oauth/token" -Method POST -Headers $authHeaders -Body $body
        $global:AccessToken = $authResponse.access_token
        $logTextBox.Text += "Connected to NinjaOne API`r`n"

        # Load devices from cache or fetch if not available
        $cacheTimestamp = Load-DeviceCache
        if (-not $cacheTimestamp) {
            # Fetch devices for lookup
            $headers = @{
                "accept" = "application/json"
                "Authorization" = "Bearer $global:AccessToken"
            }
            $devices = Invoke-RestMethod -Uri "https://$NinjaOneInstance/v2/devices" -Method GET -Headers $headers
            $global:DeviceLookup = @{}
            foreach ($device in $devices) {
                $global:DeviceLookup[$device.id] = $device.systemName
            }
            Save-DeviceCache
        }
        $logTextBox.Text += "Device lookup table ready with $($global:DeviceLookup.Count) devices`r`n"

        $connectionIndicator.Fill = "Green"
        $generateButton.IsEnabled = $true
    } catch {
        $logTextBox.Text += "Failed to connect: $_`r`n"
        $connectionIndicator.Fill = "Red"
        $generateButton.IsEnabled = $false
    }
}

# Function to Fetch Script Executions with Pagination
function Get-ScriptExecutions {
    $NinjaOneInstance = $instanceComboBox.SelectedItem.Content
    $headers = @{
        "accept" = "application/json"
        "Authorization" = "Bearer $global:AccessToken"
    }

    # Calculate date range based on selection
    $dateRange = $dateRangeFilter.SelectedItem.Content
    $afterDate = switch ($dateRange) {
        "1hr"  { (Get-Date).AddHours(-1) }
        "12hr" { (Get-Date).AddHours(-12) }
        "1day" { (Get-Date).AddDays(-1) }
        "7day" { (Get-Date).AddDays(-7) }
        "30day" { (Get-Date).AddDays(-30) }
        default { (Get-Date).AddHours(-1) }
    }
    $afterDateFormatted = $afterDate.ToString("yyyyMMdd")
    $today = Get-Date -Format "yyyyMMdd"

    # Pre-filter devices if DeviceNameFilter is set
    $deviceIds = @()
    if ($deviceNameFilter.Text) {
        $deviceIds = $global:DeviceLookup.GetEnumerator() | Where-Object { $_.Value -like "*$($deviceNameFilter.Text)*" } | ForEach-Object { $_.Key }
        if (-not $deviceIds) {
            $logTextBox.Text += "No devices match the device name filter '$($deviceNameFilter.Text)'.`r`n"
            return $null
        }
    }

    # Fetch activities (no caching for activities)
    $deviceParam = if ($deviceIds) { "&deviceId=" + ($deviceIds -join ",") } else { "" }
    $activities_url = "https://$NinjaOneInstance/api/v2/activities?class=DEVICE&type=ACTION&status=COMPLETED&after=${afterDateFormatted}&pageSize=1000${deviceParam}"
    $logTextBox.Text += "Fetching script executions with URL: $activities_url`r`n"

    $allActivities = New-Object 'System.Collections.Generic.List[PSCustomObject]'
    try {
        $activitiesRemaining = $true
        $olderThan = $null

        while ($activitiesRemaining) {
            $url = if ($olderThan) {
                "https://$NinjaOneInstance/api/v2/activities?class=DEVICE&type=ACTION&status=COMPLETED&after=${afterDateFormatted}&olderThan=${olderThan}&pageSize=1000${deviceParam}"
            } else {
                $activities_url
            }

            $response = Invoke-RestMethod -Uri $url -Method GET -Headers $headers
            $activities = $response.activities

            if ($activities.count -eq 0) {
                $activitiesRemaining = $false
            } else {
                foreach ($activity in $activities) {
                    $allActivities.Add($activity)
                }
                $olderThan = $activities[-1].id
                $logTextBox.Text += "Fetched $($activities.count) activities (total: $($allActivities.count))`r`n"
            }
        }
    } catch {
        $logTextBox.Text += "Error fetching script executions: $_`r`n"
        return $null
    }

    # Convert and format script executions
    $scriptExecutions = New-Object 'System.Collections.Generic.List[PSCustomObject]'
    $resultFilterValue = $resultFilter.SelectedItem.Content
    foreach ($activity in $allActivities) {
        if ($activity.activityType -eq "ACTION" -and $activity.sourceName -like "Run *") {
            # Apply result filter
            if ($resultFilterValue -ne "ALL" -and $activity.activityResult -ne $resultFilterValue) {
                continue
            }

            # Extract script name from sourceName (strip "Run ")
            $scriptName = if ($activity.sourceName -match "^Run\s+(.+)$") { $matches[1] } else { "Unknown Script" }

            # Apply script name filter
            if ($scriptNameFilter.Text -and $scriptName -notlike "*$($scriptNameFilter.Text)*") {
                continue
            }

            # Convert Unix timestamp to readable format (UTC) only for matching entries
            $timestamp = ([System.DateTimeOffset]::FromUnixTimeSeconds($activity.activityTime)).DateTime.ToString("yyyy-MM-dd HH:mm:ss")

            $output = if ($activity.message) { $activity.message } else { "Action: Run $scriptName" }
            
            # Check for errors in output if Result is SUCCESS
            $errorFlag = $false
            $errorMessage = $null
            if ($activity.activityResult -eq "SUCCESS") {
                $errorFlag, $errorMessage = Detect-ErrorInOutput $output
            }

            $execution = [PSCustomObject]@{
                DeviceName = $global:DeviceLookup[$activity.deviceId]
                ScriptName = $scriptName  # Renamed from ActionCompleted and stripped "Run"
                Result = $activity.activityResult
                Output = $output
                Timestamp = $timestamp
                ErrorFlag = $errorFlag
                ErrorMessage = if ($errorFlag) { "Potential error: $errorMessage" } else { $null }
            }
            $scriptExecutions.Add($execution)
        }
    }

    if ($scriptExecutions.Count -eq 0) {
        $logTextBox.Text += "No script executions found for the specified period and filters.`r`n"
        return $null
    }

    return $scriptExecutions
}

# Event Handlers
$connectButton.Add_Click({
    $logTextBox.Text = ""
    Connect-NinjaOneAPI
})

$clearCacheButton.Add_Click({
    $logTextBox.Text += "Clearing device cache...`r`n"
    if (Test-Path $deviceCacheFile) {
        try {
            Remove-Item $deviceCacheFile -ErrorAction Stop
            $logTextBox.Text += "Cleared device cache at $deviceCacheFile`r`n"
        } catch {
            $logTextBox.Text += "Failed to clear device cache at $deviceCacheFile. Error: $_`r`n"
        }
    }
    $global:DeviceLookup = @{}
    $logTextBox.Text += "Device cache cleared. Please reconnect to API to refresh.`r`n"
})

$generateButton.Add_Click({
    $scriptExecutions = Get-ScriptExecutions
    
    if ($scriptExecutions) {
        $reportDataGrid.ItemsSource = $scriptExecutions
        $exportButton.IsEnabled = $true
        $logTextBox.Text += "Script Executions Report generated successfully - $(Get-Date)`r`n"
    } else {
        $reportDataGrid.ItemsSource = $null
        $exportButton.IsEnabled = $false
    }
})

$exportButton.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.FileName = "ScriptExecutionsReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $reportDataGrid.ItemsSource | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
        $logTextBox.Text += "Report exported to $($saveFileDialog.FileName)`r`n"
    }
})

# Show the Window
$window.ShowDialog() | Out-Null
