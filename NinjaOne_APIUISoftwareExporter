<#
.SYNOPSIS
    A PowerShell script with a WPF GUI to connect to the NinjaOne API, retrieve software inventories from specified organizations, 
    filter software based on user-defined exclusions, and export results to CSV files. Additionally, it allows browsing and exporting 
    selected software across devices.

.DESCRIPTION
    This script provides a graphical interface to interact with the NinjaOne RMM platform. It authenticates using client credentials, 
    fetches organizations and software data, and allows users to:
    - Select organizations to query.
    - Exclude predefined or custom software (e.g., Microsoft Office, .NET, C++ Redistributables).
    - Choose version display options (hide, combine, or list versions).
    - Export software inventories per organization or as a master list to CSV files in a specified folder.
    - Browse and export specific software details across devices in a separate tab.
    The script uses a predefined whitelist of common Windows applications and supports importing custom exclusion lists via CSV.

.PARAMETERS
    None. Configuration is handled via the GUI, with defaults for:
    - $NinjaOneInstance: NinjaOne instance URL (e.g., app.ninjarmm.com).
    - $NinjaOneClientId: Client ID for API authentication.
    - $NinjaOneClientSecret: Client secret for API authentication.
    - $CSVOutputFolder: Output directory for CSV exports (default: C:\Temp).
    - $SoftwareWhitelist: Array of software to exclude by default (e.g., Windows Calculator, Notepad).
    - $OfficeApps: Array of Microsoft Office applications to optionally exclude.
    - $DotNetApps: Array of .NET-related applications to optionally exclude.
    - $CppRedistApps: Array of C++ Redistributables to optionally exclude.

.FUNCTIONS
    Get-NinjaOneToken
        Retrieves or refreshes an OAuth token for NinjaOne API authentication.
    Connect-NinjaOne
        Establishes a connection to the NinjaOne API using provided credentials.
    Invoke-NinjaOneRequest
        Sends HTTP requests (GET, POST, etc.) to the NinjaOne API and handles responses.

.NOTES
    - Requires the PresentationFramework assembly for WPF GUI functionality.
    - The script caches software data globally to improve performance in the "All Installed Software" tab.
    - CSV exports include fields like organization, software name, publisher, versions, device count, device IDs, device names, and product code.
    - Version display options: hide versions, combine versions into a single field, or list each version separately.
    - The GUI supports dynamic organization searching, software filtering, and exclusion list management (add, remove, import from CSV).

.EXAMPLE
    Run the script to launch the GUI:
    .\NinjaOneSoftwareExporter.ps1
    - Enter NinjaOne instance, client ID, and secret.
    - Select organizations, configure exclusions, and choose version display.
    - Click "Run Export" to generate CSV files in the specified folder.
    - Use the "All Installed Software" tab to browse and export specific software details.

.REQUIREMENTS
    - PowerShell 5.1 or later.

.OUTPUTS
    - CSV files per organization (or a master list) containing software inventory data.
    - A separate CSV file for selected software exports from the "All Installed Software" tab.
    - Log messages in the GUI for connection status, processing, and errors.

**Disclaimer**: This script is provided "as is" under the MIT License. 
    Use at your own risk; the author is not responsible for any damages or issues arising from its use. 
    Licensed under MIT: Copyright (c) Spencer Heath. Permission is granted to use, copy, modify, and distribute this software freely, 
    provided the original copyright and this notice are retained. See https://opensource.org/licenses/MIT for full details.
#>

Add-Type -AssemblyName PresentationFramework -ErrorAction Stop

$NinjaOneInstance = ''
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''
$CSVOutputFolder = "C:\Temp"
$SoftwareWhitelist = @(
    "Microsoft.WindowsAlarms",              # Alarms & Clock
    "Microsoft.WindowsCalculator",          # Calculator
    "Microsoft.WindowsCamera",              # Camera
    "Microsoft.windowscommunicationsapps",  # Mail and Calendar
    "Microsoft.WindowsFeedbackHub",         # Feedback Hub
    "Microsoft.WindowsMaps",                # Maps
    "Microsoft.Windows.Photos",             # Photos
    "Microsoft.WindowsSoundRecorder",       # Voice Recorder
    "Microsoft.YourPhone",                  # Your Phone
    "Microsoft.ZuneMusic",                  # Groove Music
    "Microsoft.ZuneVideo",                  # Movies & TV
    "Microsoft.People",                     # People
    "Microsoft.SkypeApp",                   # Skype (pre-installed)
    "Microsoft.MixedReality.Portal",        # Mixed Reality Portal
    "Microsoft.GetHelp",                    # Get Help
    "Microsoft.Getstarted",                 # Tips
    "Microsoft.Microsoft3DViewer",          # 3D Viewer
    "Microsoft.MicrosoftSolitaireCollection", # Solitaire Collection
    "Microsoft.MicrosoftStickyNotes",       # Sticky Notes
    "Microsoft.MSPaint",                    # Paint 3D
    "Microsoft.Office.OneNote",             # OneNote (pre-installed version)
    "Microsoft.XboxIdentityProvider",       # Xbox Identity Provider
    "Microsoft.VP9VideoExtensions",         # VP9 Video Extensions
    "Microsoft.WebMediaExtensions",         # Web Media Extensions
    "Microsoft.WebpImageExtension",         # WebP Image Extension
    "Microsoft.ScreenSketch",               # Snip & Sketch
    "Microsoft.WindowsStore",               # Microsoft Store
    "Microsoft.Xbox.TCUI",                  # Xbox TCUI
    "Microsoft.XboxGameOverlay",            # Xbox Game Overlay
    "Microsoft.XboxGamingOverlay",          # Xbox Gaming Overlay
    "Microsoft.StorePurchaseApp",           # Store Purchase App
    "Microsoft.DesktopAppInstaller",        # Desktop App Installer
    "Microsoft.Windows.DevHome",            # Dev Home
    "Microsoft.MPEG2VideoExtension",        # MPEG-2 Video Extension
    "Microsoft Update Health Tools",        # Update Health Tools
    "Microsoft.HEVCVideoExtension",         # HEVC Video Extension
    "MicrosoftWindows.CrossDevice",         # Cross Device
    "Microsoft.SecHealthUI",                # Windows Security
    "Microsoft.WidgetsPlatformRuntime",     # Widgets Platform
    "Microsoft.WindowsNotepad",             # Notepad
    "Microsoft.RawImageExtension",          # Raw Image Extension
    "Microsoft.Paint",                      # Paint
    "Microsoft.Copilot",                    # Copilot
    "Microsoft.OneDriveSync",               # OneDrive Sync
    "Microsoft.OutlookForWindows",          # Outlook for Windows
    "Microsoft.OneDrive",                   # OneDrive
    "Microsoft.Todos",                      # To Do
    "Microsoft.RemoteDesktop",              # Remote Desktop
    "Microsoft.BingNews",                   # Bing News
    "Microsoft.MicrosoftEdge.Stable",       # Edge Stable
    "Microsoft.MicrosoftOfficeHub",         # Office Hub
    "MicrosoftWindows.Client.WebExperience", # Web Experience Pack
    "Microsoft.BingWeather",                # Bing Weather
    "Microsoft.549981C3F5F10",              # Cortana
    "Microsoft.GamingApp",                  # Gaming App
    "Clipchamp.Clipchamp",                  # Clipchamp
    "7EE7776C.LinkedInforWindows",          # LinkedIn for Windows
    "Microsoft.ApplicationCompatibilityEnhancements", # Application Compatibility Enhancements
    "Microsoft.AV1VideoExtension",          # AV1 Video Extension
    "Microsoft.AVCEncoderVideoExtension",   # AVC Encoder Video Extension
    "Microsoft.BingSearch",                 # Bing Search
    "Microsoft.BingTranslator",             # Bing Translator
    "Microsoft.D3DMappingLayers",           # Direct3D Mapping Layers
    "Microsoft.HEIFImageExtension",         # HEIF Image Extension
    "Microsoft.Ink.Handwriting.Main.en-US.1.0.1", # Ink Handwriting Support (English US)
    "Microsoft.Messaging",                  # Messaging App
    "Microsoft.MicrosoftJournal",           # Microsoft Journal
    "Microsoft.MicrosoftPowerBIForWindows", # Power BI for Windows
    "Microsoft.MinecraftEducationEdition",  # Minecraft Education Edition
    "Microsoft.NetworkSpeedTest",           # Network Speed Test
    "Microsoft.Office.Excel",               # Excel (pre-installed version)
    "Microsoft.Office.OneNoteVirtualPrinter", # OneNote Virtual Printer
    "Microsoft.Office.Sway",                # Sway
    "Microsoft.OfficePushNotificationUtility", # Office Push Notification Utility
    "Microsoft.OneConnect",                 # OneConnect (deprecated mobile connectivity app)
    "Microsoft.OfficeLens",                 # Office Lens
    "Microsoft.PowerAutomateDesktop",       # Power Automate Desktop
    "Microsoft.PowerToys.FileLocksmithContextMenu", # PowerToys File Locksmith Context Menu
    "Microsoft.PowerToys.ImageResizerContextMenu", # PowerToys Image Resizer Context Menu
    "Microsoft.PowerToys.PowerRenameContextMenu", # PowerToys Power Rename Context Menu
    "Microsoft.Print3D",                    # Print 3D
    "Microsoft.Studios.Wordament",          # Wordament
    "Microsoft.Wallet",                     # Microsoft Wallet
    "Microsoft.Whiteboard",                 # Whiteboard
    "Microsoft.WinAppRuntime.DDLM.2000.609.1413.0-x6-p1", # Windows App Runtime DDLM (x64, version-specific)
    "Microsoft.WinAppRuntime.DDLM.2000.609.1413.0-x8-p1", # Windows App Runtime DDLM (x86, version-specific)
    "Microsoft.WinAppRuntime.DDLM.4000.1082.2259.0-x6", # Windows App Runtime DDLM (x64, version-specific)
    "Microsoft.WinAppRuntime.DDLM.4000.1082.2259.0-x8", # Windows App Runtime DDLM (x86, version-specific)
    "Microsoft.Windows.Ai.Copilot.Provider", # Copilot AI Provider
    "Microsoft.WindowsScan",                # Windows Scan
    "Microsoft.WindowsTerminal",            # Windows Terminal
    "Microsoft.Winget.Platform.Source",     # Winget Platform Source   
    "Microsoft.Winget.Source"               # Winget Source
)

$OfficeApps = @(
    "Microsoft.Office.Word",
    "Microsoft.Office.Excel",
    "Microsoft.Office.PowerPoint",
    "Microsoft.Office.Outlook",
    "Microsoft.Office.Access",
    "Microsoft.Office.Publisher",
    "Microsoft.Office.OneDrive",
    "Microsoft 365 Apps for business - en-us", # English (US)
    "Microsoft 365 Apps for business - fr-fr", # French (France)
    "Microsoft 365 Apps for business - es-es", # Spanish (Spain)
    "Microsoft 365 Apps for business - de-de", # German (Germany)
    "Microsoft 365 Apps for business - it-it", # Italian (Italy)
    "Microsoft 365 Apps for business - ja-jp", # Japanese (Japan)
    "Microsoft 365 Apps for business - zh-cn", # Chinese (Simplified)
    "Microsoft 365 Apps for business - zh-tw", # Chinese (Traditional)
    "Microsoft 365 Apps for business - pt-br", # Portuguese (Brazil)
    "Microsoft 365 Apps for business - ru-ru", # Russian
    "Microsoft 365 Apps for business - ko-kr", # Korean
    "Microsoft 365 Apps for business - nl-nl", # Dutch (Netherlands)
    "Microsoft 365 Apps for business - sv-se"  # Swedish (Sweden)
)

$DotNetApps = @(
    "Microsoft.NET",
    "Microsoft .NET Framework",
    "Microsoft .NET Core"
)

$CppRedistApps = @(
    "Microsoft Visual C++",
    "Microsoft C++ Redistributable"
)

# Global cache for software list
$global:CachedSoftware = @()

function Get-NinjaOneToken {
    [CmdletBinding()]
    param()

    if ($Script:NinjaOneInstance -and $Script:NinjaOneClientID -and $Script:NinjaOneClientSecret) {
        if ($Script:NinjaTokenExpiry -and (Get-Date) -lt $Script:NinjaTokenExpiry) {
            return $Script:NinjaToken
        }
        else {
            if ($Script:NinjaOneRefreshToken) {
                $Body = @{
                    'grant_type'    = 'refresh_token'
                    'client_id'     = $Script:NinjaOneClientID
                    'client_secret' = $Script:NinjaOneClientSecret
                    'refresh_token' = $Script:NinjaOneRefreshToken
                }
            }
            else {
                $body = @{
                    grant_type    = 'client_credentials'
                    client_id     = $Script:NinjaOneClientID
                    client_secret = $Script:NinjaOneClientSecret
                    scope         = 'monitoring management'
                }
            }

            $token = Invoke-RestMethod -Uri "https://$($Script:NinjaOneInstance -replace '/ws','')/ws/oauth/token" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
    
            $Script:NinjaTokenExpiry = (Get-Date).AddSeconds($Token.expires_in)
            $Script:NinjaToken = $token
            
            $logTextBox.Text += "Fetched New Token`r`n"
            return $token
        }
    }
    else {
        Throw 'Please run Connect-NinjaOne first'
    }
}

function Connect-NinjaOne {
    [CmdletBinding()]
    param (
        [Parameter(mandatory = $true)]
        $NinjaOneInstance,
        [Parameter(mandatory = $true)]
        $NinjaOneClientID,
        [Parameter(mandatory = $true)]
        $NinjaOneClientSecret,
        $NinjaOneRefreshToken
    )

    $Script:NinjaOneInstance = $NinjaOneInstance
    $Script:NinjaOneClientID = $NinjaOneClientID
    $Script:NinjaOneClientSecret = $NinjaOneClientSecret
    $Script:NinjaOneRefreshToken = $NinjaOneRefreshToken
    
    try {
        $Null = Get-NinjaOneToken -ea Stop
    }
    catch {
        Throw "Failed to Connect to NinjaOne: $_"
    }
}

function Invoke-NinjaOneRequest {
    param(
        $Method,
        $Body,
        $InputObject,
        $Path,
        $QueryParams,
        [Switch]$Paginate,
        [Switch]$AsArray
    )

    $Token = Get-NinjaOneToken

    if ($InputObject) {
        if ($AsArray) {
            $Body = $InputObject | ConvertTo-Json -depth 100
            if (($InputObject | Measure-Object).count -eq 1 ) {
                $Body = '[' + $Body + ']'
            }
        }
        else {
            $Body = $InputObject | ConvertTo-Json -depth 100
        }
    }

    try {
        if ($Method -in @('GET', 'DELETE')) {
            if ($Paginate) {
                $After = 0
                $PageSize = 1000
                $NinjaResult = do {
                    $Result = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)?pageSize=$PageSize&after=$After$(if ($QueryParams){"&$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json' -UseBasicParsing
                    $Result
                    $ResultCount = ($Result.id | Measure-Object -Maximum)
                    $After = $ResultCount.maximum
                } while ($ResultCount.count -eq $PageSize)
            }
            else {
                $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json; charset=utf-8' -UseBasicParsing
            }
        }
        elseif ($Method -in @('PATCH', 'PUT', 'POST')) {
            $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -Body $Body -ContentType 'application/json; charset=utf-8' -UseBasicParsing
        }
        else {
            Throw 'Unknown Method'
        }
    }
    catch {
        Throw "Error Occurred: $_"
    }

    try {
        return $NinjaResult.content | ConvertFrom-Json -ea stop
    }
    catch {
        return $NinjaResult.content
    }
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="NinjaOne Software Inventory Exporter" Height="850" Width="1000" WindowStartupLocation="CenterScreen" Background="#1E1E1E">
    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="300"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar for Exclude List -->
        <Border Grid.Column="0" Background="#2D2D2D" Margin="0,0,15,0" BorderBrush="#444444" BorderThickness="1" Padding="10">
            <StackPanel>
                <Label Content="Exclude Software List" Foreground="White" FontSize="16" Margin="0,0,0,10"/>
                <ListBox x:Name="ExcludeListBox" Height="500" Background="#2D2D2D" Foreground="White" BorderBrush="#444444" Margin="0,0,0,15"/>
                <Button x:Name="AddExcludeButton" Content="Add" Width="120" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Button x:Name="RemoveExcludeButton" Content="Remove" Width="120" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Button x:Name="RemoveAllExcludeButton" Content="Remove All" Width="120" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                <Button x:Name="ImportExcludeButton" Content="Import CSV" Width="120" Height="30" Margin="0,0,0,10" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
            </StackPanel>
        </Border>

        <!-- Tab Control for Main Content -->
        <TabControl Grid.Column="1" Background="#1E1E1E" BorderBrush="#444444">
            <TabItem Header="Software Export" Background="#2D2D2D" Foreground="White">
                <Grid Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Label Grid.Row="0" Content="NinjaOne Instance:" Foreground="White" FontSize="14" Margin="0,0,0,10"/>
                    <ComboBox x:Name="InstanceComboBox" Grid.Row="0" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" SelectedIndex="0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555">
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
                                                <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
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

                    <Label Grid.Row="1" Content="Client ID:" Foreground="White" FontSize="14" Margin="0,0,0,10"/>
                    <TextBox x:Name="ClientIdTextBox" Grid.Row="1" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>

                    <Label Grid.Row="2" Content="Client Secret:" Foreground="White" FontSize="14" Margin="0,0,0,10"/>
                    <TextBox x:Name="ClientSecretTextBox" Grid.Row="2" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>

                    <Button x:Name="ConnectButton" Grid.Row="3" Content="Connect to API" Width="120" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                    <Ellipse x:Name="ConnectionIndicator" Grid.Row="3" Width="15" Height="15" Margin="280,0,0,15" HorizontalAlignment="Left" Fill="Red"/>

                    <Label Grid.Row="4" Content="Output Folder:" Foreground="White" FontSize="14" Margin="0,0,0,10"/>
                    <TextBox x:Name="OutputFolderTextBox" Grid.Row="4" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Text="$CSVOutputFolder" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>

                    <Label Grid.Row="5" Content="Organizations:" Foreground="White" FontSize="14" Margin="0,0,0,10"/>
                    <TextBox x:Name="OrgSearchBox" Grid.Row="5" Width="250" Height="30" Margin="150,0,0,5" HorizontalAlignment="Left" Text="Search organizations..." Foreground="Gray" Background="#3C3C3C" BorderBrush="#555555"/>
                    <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="410,0,0,5" HorizontalAlignment="Left">
                        <Button x:Name="CheckAllButton" Content="Check All" Width="100" Height="30" Margin="0,0,15,0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                        <Button x:Name="UncheckAllButton" Content="Uncheck All" Width="100" Height="30" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                    </StackPanel>
                    <ListBox x:Name="OrgsListBox" Grid.Row="6" Height="250" Margin="0,0,0,15" SelectionMode="Multiple" Background="#3C3C3C" Foreground="White" BorderBrush="#555555">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <CheckBox Content="{Binding Name}" IsChecked="{Binding IsSelected, RelativeSource={RelativeSource AncestorType={x:Type ListBoxItem}}}" Foreground="White"/>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>

                    <Grid Grid.Row="7" Margin="0,0,0,15">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <CheckBox x:Name="ExcludeOfficeCheckBox" Grid.Column="0" Content="Exclude Office Apps" Margin="0,0,20,0" Foreground="White" Background="#3C3C3C"/>
                        <CheckBox x:Name="ExcludeDotNetCheckBox" Grid.Column="1" Content="Exclude .NET" Margin="0,0,20,0" Foreground="White" Background="#3C3C3C"/>
                        <CheckBox x:Name="ExcludeCppCheckBox" Grid.Column="2" Content="Exclude C++ Redist" Margin="0,0,20,0" Foreground="White" Background="#3C3C3C"/>
                        <Label Grid.Column="3" Content="Version Display:" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <ComboBox x:Name="VersionDisplayComboBox" Grid.Column="4" Width="200" Height="30" SelectedIndex="0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555">
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
                                                    <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
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
                            <ComboBoxItem Content="Hide Versions"/>
                            <ComboBoxItem Content="Combine Versions"/>
                            <ComboBoxItem Content="List Each Version"/>
                        </ComboBox>
                    </Grid>

                    <Label Grid.Row="8" Content="Log:" Foreground="White" FontSize="14" Margin="0,0,0,5"/>
                    <TextBox x:Name="LogTextBox" Grid.Row="8" Height="250" Margin="0,0,0,15" AcceptsReturn="True" VerticalScrollBarVisibility="Visible" IsReadOnly="True" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                    <Button x:Name="RunButton" Grid.Row="9" Content="Run Export" Width="120" Height="30" Margin="0,0,0,0" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" IsEnabled="False"/>
                </Grid>
            </TabItem>

            <TabItem Header="All Installed Software" Background="#2D2D2D" Foreground="White">
                <Grid Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Label Grid.Row="0" Content="Search Software:" Foreground="White" FontSize="14" Margin="0,0,0,10"/>
                    <TextBox x:Name="SoftwareSearchBox" Grid.Row="0" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Text="Search software..." Foreground="Gray" Background="#3C3C3C" BorderBrush="#555555"/>

                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="SoftwareCheckAllButton" Content="Check All" Width="100" Height="30" Margin="0,0,15,0" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                        <Button x:Name="SoftwareUncheckAllButton" Content="Uncheck All" Width="100" Height="30" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>
                    </StackPanel>

                    <ListBox x:Name="SoftwareListBox" Grid.Row="2" Height="500" SelectionMode="Multiple" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" Margin="0,0,0,15">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <CheckBox Content="{Binding DisplayName}" IsChecked="{Binding IsSelected, RelativeSource={RelativeSource AncestorType={x:Type ListBoxItem}}}" Foreground="White"/>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>

                    <Label Grid.Row="3" Content="Log:" Foreground="White" FontSize="14" Margin="0,0,0,5"/>
                    <TextBox x:Name="SoftwareLogTextBox" Grid.Row="3" Height="150" Margin="0,0,0,15" AcceptsReturn="True" VerticalScrollBarVisibility="Visible" IsReadOnly="True" Background="#3C3C3C" Foreground="White" BorderBrush="#555555"/>

                    <Button x:Name="ExportSoftwareButton" Grid.Row="4" Content="Export Selected" Width="120" Height="30" Margin="0,0,0,0" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" IsEnabled="False"/>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

try {
    $window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
}
catch {
    Write-Error "Failed to load XAML: $_"
    exit 1
}

$instanceComboBox = $window.FindName("InstanceComboBox")
$clientIdTextBox = $window.FindName("ClientIdTextBox")
$clientSecretTextBox = $window.FindName("ClientSecretTextBox")
$connectButton = $window.FindName("ConnectButton")
$connectionIndicator = $window.FindName("ConnectionIndicator")
$outputFolderTextBox = $window.FindName("OutputFolderTextBox")
$orgSearchBox = $window.FindName("OrgSearchBox")
$checkAllButton = $window.FindName("CheckAllButton")
$uncheckAllButton = $window.FindName("UncheckAllButton")
$orgsListBox = $window.FindName("OrgsListBox")
$logTextBox = $window.FindName("LogTextBox")
$runButton = $window.FindName("RunButton")
$excludeListBox = $window.FindName("ExcludeListBox")
$addExcludeButton = $window.FindName("AddExcludeButton")
$removeExcludeButton = $window.FindName("RemoveExcludeButton")
$removeAllExcludeButton = $window.FindName("RemoveAllExcludeButton")
$importExcludeButton = $window.FindName("ImportExcludeButton")
$excludeOfficeCheckBox = $window.FindName("ExcludeOfficeCheckBox")
$excludeDotNetCheckBox = $window.FindName("ExcludeDotNetCheckBox")
$excludeCppCheckBox = $window.FindName("ExcludeCppCheckBox")
$versionDisplayComboBox = $window.FindName("VersionDisplayComboBox")
$softwareSearchBox = $window.FindName("SoftwareSearchBox")
$softwareListBox = $window.FindName("SoftwareListBox")
$softwareCheckAllButton = $window.FindName("SoftwareCheckAllButton")
$softwareUncheckAllButton = $window.FindName("SoftwareUncheckAllButton")
$softwareLogTextBox = $window.FindName("SoftwareLogTextBox")
$exportSoftwareButton = $window.FindName("ExportSoftwareButton")

if (-not $window -or -not $instanceComboBox -or -not $clientIdTextBox -or -not $clientSecretTextBox -or -not $connectButton -or -not $connectionIndicator -or -not $outputFolderTextBox -or -not $orgSearchBox -or -not $checkAllButton -or -not $uncheckAllButton -or -not $orgsListBox -or -not $logTextBox -or -not $runButton -or -not $excludeListBox -or -not $addExcludeButton -or -not $removeExcludeButton -or -not $removeAllExcludeButton -or -not $importExcludeButton -or -not $excludeOfficeCheckBox -or -not $excludeDotNetCheckBox -or -not $excludeCppCheckBox -or -not $versionDisplayComboBox -or -not $softwareSearchBox -or -not $softwareListBox -or -not $softwareCheckAllButton -or -not $softwareUncheckAllButton -or -not $softwareLogTextBox -or -not $exportSoftwareButton) {
    Write-Error "One or more UI elements could not be found. Check XAML naming and structure."
    exit 1
}

foreach ($item in $SoftwareWhitelist) {
    $excludeListBox.Items.Add($item) | Out-Null
}

$connectButton.Add_Click({
    $logTextBox.Text = ""
    $NinjaOneInstance = $instanceComboBox.SelectedItem.Content
    $NinjaOneClientId = $clientIdTextBox.Text
    $NinjaOneClientSecret = $clientSecretTextBox.Text

    try {
        Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
        $logTextBox.Text += "Connected to NinjaOne API`r`n"
        $connectionIndicator.Fill = "Green"

        $AllOrgs = Invoke-NinjaOneRequest -Method Get -Path "organizations"
        $logTextBox.Text += "Retrieved organizations`r`n"
        $orgsListBox.Items.Clear()
        foreach ($org in $AllOrgs) {
            $orgsListBox.Items.Add([PSCustomObject]@{Name = $org.name})
        }

        $softwareLogTextBox.Text = "Retrieving software list...`r`n"
        $AllSoftware = Invoke-NinjaOneRequest -Method Get -Path "queries/software"
        $global:CachedSoftware = $AllSoftware.results | Sort-Object name | Group-Object name | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.Name
                FullData    = $_.Group
            }
        }
        $softwareListBox.Items.Clear()
        foreach ($software in $global:CachedSoftware) {
            $softwareListBox.Items.Add($software)
        }
        $softwareLogTextBox.Text += "Retrieved $($global:CachedSoftware.Count) unique software items`r`n"
        $runButton.IsEnabled = $true
        $exportSoftwareButton.IsEnabled = $true
    }
    catch {
        $logTextBox.Text += "Failed to connect to NinjaOne API: $($_)`r`n"
        $softwareLogTextBox.Text += "Failed to connect to NinjaOne API: $($_)`r`n"
        $connectionIndicator.Fill = "Red"
        $runButton.IsEnabled = $false
        $exportSoftwareButton.IsEnabled = $false
    }
})

$orgSearchBox.Add_GotFocus({
    if ($orgSearchBox.Text -eq "Search organizations...") {
        $orgSearchBox.Text = ""
        $orgSearchBox.Foreground = "White"
    }
})

$orgSearchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($orgSearchBox.Text)) {
        $orgSearchBox.Text = "Search organizations..."
        $orgSearchBox.Foreground = "Gray"
    }
})

$orgSearchBox.Add_TextChanged({
    if ($orgSearchBox.Text -ne "Search organizations...") {
        $searchText = $orgSearchBox.Text.ToLower()
        $orgsListBox.Items.Clear()
        $AllOrgs = Invoke-NinjaOneRequest -Method Get -Path "organizations"
        foreach ($org in $AllOrgs) {
            if ($org.name.ToLower() -like "*$searchText*") {
                $orgsListBox.Items.Add([PSCustomObject]@{Name = $org.name})
            }
        }
    }
})

$checkAllButton.Add_Click({
    foreach ($item in $orgsListBox.Items) {
        $orgsListBox.SelectedItems.Add($item)
    }
})

$uncheckAllButton.Add_Click({
    $orgsListBox.SelectedItems.Clear()
})

$addExcludeButton.Add_Click({
    Add-Type -AssemblyName Microsoft.VisualBasic
    $input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Software to Exclude", "Add Exclusion")
    if ($input -and -not $excludeListBox.Items.Contains($input)) {
        $excludeListBox.Items.Add($input)
    }
})

$removeExcludeButton.Add_Click({
    if ($excludeListBox.SelectedItem) {
        $excludeListBox.Items.Remove($excludeListBox.SelectedItem)
    }
})

$removeAllExcludeButton.Add_Click({
    $excludeListBox.Items.Clear()
})

$importExcludeButton.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Select a CSV File with Software to Exclude"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $csvContent = Import-Csv -Path $openFileDialog.FileName
        foreach ($item in $csvContent) {
            $softwareName = $item.PSObject.Properties.Value | Where-Object { $_ } | Select-Object -First 1
            if ($softwareName -and -not $excludeListBox.Items.Contains($softwareName)) {
                $excludeListBox.Items.Add($softwareName)
            }
        }
        $logTextBox.Text += "Imported exclude list from $($openFileDialog.FileName)`r`n"
    }
})

$runButton.Add_Click({
    $logTextBox.Text = ""
    $CSVOutputFolder = $outputFolderTextBox.Text
    $excludeList = $excludeListBox.Items | ForEach-Object { $_ }
    if ($excludeOfficeCheckBox.IsChecked) {
        $excludeList += $OfficeApps
    }
    if ($excludeDotNetCheckBox.IsChecked) {
        $excludeList += $DotNetApps
    }
    if ($excludeCppCheckBox.IsChecked) {
        $excludeList += $CppRedistApps
    }

    if ($orgsListBox.SelectedItems.Count -eq 0) {
        $AllOrgs = Invoke-NinjaOneRequest -Method Get -Path "organizations"
        $OrgList = $AllOrgs
        $logTextBox.Text += "Processing all organizations`r`n"
    }
    else {
        $AllOrgs = Invoke-NinjaOneRequest -Method Get -Path "organizations"
        $OrgList = $AllOrgs | Where-Object { $orgsListBox.SelectedItems.Name -contains $_.name }
        if (-not $OrgList) {
            $logTextBox.Text += "No matching organizations selected`r`n"
            return
        }
    }

    $MasterList = @()

    foreach ($Org in $OrgList) {
        $OrgID = $Org.id
        $OrgName = $Org.name
        $logTextBox.Text += "Processing $($OrgName)`r`n"

        try {
            $SoftwarePath = "queries/software?df=org=$($OrgID)"
            $AllSoftwareForOrg = Invoke-NinjaOneRequest -Method Get -Path $SoftwarePath
        }
        catch {
            $logTextBox.Text += "Failed to retrieve software for $($OrgName): $($_)`r`n"
            continue
        }

        if (-not $AllSoftwareForOrg -or -not $AllSoftwareForOrg.results) {
            $logTextBox.Text += "No software results for $($OrgName)`r`n"
            continue
        }

        try {
            $DevicesPath = "devices?df=org=$($OrgID)"
            $AllDevicesForOrg = Invoke-NinjaOneRequest -Method Get -Path $DevicesPath
        }
        catch {
            $logTextBox.Text += "Failed to retrieve devices for $($OrgName): $($_)`r`n"
            continue
        }

        $deviceLookup = @{}
        foreach ($device in $AllDevicesForOrg) {
            $deviceLookup[$device.id] = $device.systemName
        }

        $filteredSoftware = $AllSoftwareForOrg.results | Where-Object { $excludeList -notcontains $_.name }

        switch ($versionDisplayComboBox.SelectedIndex) {
            0 { 
                $groupedResults = $filteredSoftware | Group-Object -Property name, publisher, productCode | ForEach-Object {
                    $deviceIds = $_.Group | ForEach-Object { $_.deviceId }
                    $deviceNames = $deviceIds | ForEach-Object { if ($deviceLookup.ContainsKey($_)) { $deviceLookup[$_] } }
                    
                    [PSCustomObject]@{
                        Organization = $OrgName
                        name         = $_.Group[0].name
                        publisher    = $_.Group[0].publisher
                        versions     = $null
                        deviceCount  = $_.Count
                        deviceId     = $deviceIds -join ","
                        deviceNames  = $deviceNames -join ","
                        productCode  = $_.Group[0].productCode
                    }
                }
            }
            1 {
                $groupedResults = $filteredSoftware | Group-Object -Property name, publisher | ForEach-Object {
                    $versions = $_.Group | Select-Object -ExpandProperty version -Unique | Sort-Object | Where-Object { $_ } | Join-String -Separator ", "
                    $deviceIds = $_.Group | ForEach-Object { $_.deviceId }
                    $deviceNames = $deviceIds | ForEach-Object { if ($deviceLookup.ContainsKey($_)) { $deviceLookup[$_] } }
                    
                    [PSCustomObject]@{
                        Organization = $OrgName
                        name         = $_.Group[0].name
                        publisher    = $_.Group[0].publisher
                        versions     = $versions
                        deviceCount  = $_.Count
                        deviceId     = $deviceIds -join ","
                        deviceNames  = $deviceNames -join ","
                        productCode  = $_.Group[0].productCode
                    }
                }
            }
            2 {
                $groupedResults = $filteredSoftware | ForEach-Object {
                    $deviceId = $_.deviceId
                    $deviceName = if ($deviceLookup.ContainsKey($deviceId)) { $deviceLookup[$deviceId] } else { "" }
                    
                    [PSCustomObject]@{
                        Organization = $OrgName
                        name         = $_.name
                        publisher    = $_.publisher
                        versions     = $_.version
                        deviceCount  = 1
                        deviceId     = $deviceId
                        deviceNames  = $deviceName
                        productCode  = $_.productCode
                    }
                }
            }
        }

        $finalResults = $groupedResults | Sort-Object -Property deviceCount -Descending
        $MasterList += $finalResults

        $timestamp = Get-Date -Format "yyyy-MM-dd HH-mm"
        $filename = "Software Inventory Export - $($timestamp) - $($OrgName).csv"
        $filepath = Join-Path -Path $CSVOutputFolder -ChildPath $filename

        $finalResults | Export-Csv -Path $filepath -NoTypeInformation
        $logTextBox.Text += "CSV exported for $($OrgName): $($filepath)`r`n"
    }

    if ($orgsListBox.SelectedItems.Count -eq 0 -and $MasterList) {
        $masterFilename = "Software Inventory Export - $($timestamp) - AllOrgs.csv"
        $masterFilepath = Join-Path -Path $CSVOutputFolder -ChildPath $masterFilename
        $MasterList | Export-Csv -Path $masterFilepath -NoTypeInformation
        $logTextBox.Text += "Master list exported: $($masterFilepath)`r`n"
    }
    elseif (-not $MasterList) {
        $logTextBox.Text += "No data collected to export.`r`n"
    }
})

$softwareSearchBox.Add_GotFocus({
    if ($softwareSearchBox.Text -eq "Search software...") {
        $softwareSearchBox.Text = ""
        $softwareSearchBox.Foreground = "White"
    }
})

$softwareSearchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($softwareSearchBox.Text)) {
        $softwareSearchBox.Text = "Search software..."
        $softwareSearchBox.Foreground = "Gray"
    }
})

$softwareSearchBox.Add_TextChanged({
    if ($softwareSearchBox.Text -ne "Search software...") {
        $searchText = $softwareSearchBox.Text.ToLower()
        $softwareListBox.Items.Clear()
        
        $filteredSoftware = $global:CachedSoftware | Where-Object { $_.DisplayName.ToLower() -like "*$searchText*" }
        
        foreach ($software in $filteredSoftware) {
            $softwareListBox.Items.Add($software)
        }
    }
})

$softwareCheckAllButton.Add_Click({
    foreach ($item in $softwareListBox.Items) {
        $softwareListBox.SelectedItems.Add($item)
    }
})

$softwareUncheckAllButton.Add_Click({
    $softwareListBox.SelectedItems.Clear()
})

$exportSoftwareButton.Add_Click({
    $softwareLogTextBox.Text = ""
    $CSVOutputFolder = $outputFolderTextBox.Text
    $selectedSoftware = $softwareListBox.SelectedItems
    
    if ($selectedSoftware.Count -eq 0) {
        $softwareLogTextBox.Text += "No software selected for export`r`n"
        return
    }

    $exportData = @()
    $AllDevices = Invoke-NinjaOneRequest -Method Get -Path "devices"
    $deviceLookup = @{}
    foreach ($device in $AllDevices) {
        $deviceLookup[$device.id] = $device.systemName
    }

    foreach ($software in $selectedSoftware) {
        $softwareLogTextBox.Text += "Processing $($software.DisplayName)`r`n"
        foreach ($instance in $software.FullData) {
            $deviceName = $deviceLookup[$instance.deviceId] ?? "Unknown Device"
            $exportData += [PSCustomObject]@{
                SoftwareName = $software.DisplayName
                Version      = $instance.version
                Publisher    = $instance.publisher
                DeviceName   = $deviceName
                DeviceId     = $instance.deviceId
                ProductCode  = $instance.productCode
            }
        }
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH-mm"
    $filename = "Selected Software Export - $($timestamp).csv"
    $filepath = Join-Path -Path $CSVOutputFolder -ChildPath $filename
    
    $exportData | Export-Csv -Path $filepath -NoTypeInformation
    $softwareLogTextBox.Text += "Exported $($selectedSoftware.Count) software items to: $($filepath)`r`n"
})

$window.ShowDialog() | Out-Null
