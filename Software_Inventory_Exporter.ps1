<#
.SYNOPSIS
    A PowerShell script with a WPF GUI to connect to the NinjaOne API, retrieve software inventories from specified organizations, 
    filter software based on user-defined exclusions, and export results to CSV files asynchronously. It also allows browsing and exporting 
    selected software across devices.

.DESCRIPTION
    This script provides a graphical interface to interact with the NinjaOne RMM platform. It uses async operations to authenticate, 
    fetch organizations and software data, and allows users to:
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
    - $NinjaOneClientSecret: Client secret for API authentication (not stored in config file).
    - $CSVOutputFolder: Output directory for CSV exports (default: C:\Temp).
    - $SoftwareWhitelist: Array of software to exclude by default (e.g., Windows Calculator, Notepad).
    - $OfficeApps: Array of Microsoft Office applications to optionally exclude.
    - $DotNetApps: Array of .NET-related applications to optionally exclude.
    - $CppRedistApps: Array of C++ Redistributables to optionally exclude.

.FUNCTIONS
    Get-NinjaOneToken
        Retrieves or refreshes an OAuth token for NinjaOne API authentication asynchronously.
    Connect-NinjaOne
        Establishes a connection to the NinjaOne API using provided credentials.
    Invoke-NinjaOneRequestAsync
        Sends HTTP requests (GET, POST, etc.) to the NinjaOne API asynchronously and handles responses.

.EXAMPLE
    Run the script to launch the GUI:
    .\NinjaOneSoftwareExporter_Async.ps1
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
Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
Add-Type -AssemblyName System.Net.Http


$httpClientHandler = New-Object System.Net.Http.HttpClientHandler
$httpClient = New-Object System.Net.Http.HttpClient($httpClientHandler)
$httpClient.Timeout = [System.TimeSpan]::FromSeconds(30)


$global:NinjaOneInstance = ''
$global:NinjaOneClientId = ''
$global:NinjaOneClientSecret = ''
$global:NinjaToken = $null
$global:NinjaTokenExpiry = $null

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
    "Microsoft Moderator",                  # Microsoft Moderator
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

function Convert-SizeToReadable {
    param (
        [Parameter(Mandatory=$true)]
        [double]$Bytes
    )

    if ($Bytes -ge 1GB) {
        "{0:N3} GB" -f ($Bytes / 1GB)
    }
    elseif ($Bytes -ge 1MB) {
        "{0:N3} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -ge 1KB) {
        "{0:N3} KB" -f ($Bytes / 1KB)
    }
    else {
        "$Bytes Bytes"
    }
}


function Convert-UnixTimestampToDateTime {
    param (
        [Parameter(Mandatory=$true)]
        [double]$UnixTimestamp
    )

    $epoch = [DateTime]::Parse("1970-01-01T00:00:00Z")
    $dateTime = $epoch.AddSeconds($UnixTimestamp).ToLocalTime()
    $dateTime.ToString("MM/dd/yyyy hh:mm:ss tt")
}


$global:CachedSoftware = @()

function Get-NinjaOneToken {
    [CmdletBinding()]
    param()

    if ($global:NinjaOneInstance -and $global:NinjaOneClientId -and $global:NinjaOneClientSecret) {
        if ($global:NinjaTokenExpiry -and (Get-Date) -lt $global:NinjaTokenExpiry) {
            return $global:NinjaToken
        }
        else {
         
            $dict = New-Object 'System.Collections.Generic.Dictionary[string,string]'

            if ($Script:NinjaOneRefreshToken) {
                $dict.Add('grant_type', 'refresh_token')
                $dict.Add('client_id', $global:NinjaOneClientId)
                $dict.Add('client_secret', $global:NinjaOneClientSecret)
                $dict.Add('refresh_token', $Script:NinjaOneRefreshToken)
            }
            else {
                $dict.Add('grant_type', 'client_credentials')
                $dict.Add('client_id', $global:NinjaOneClientId)
                $dict.Add('client_secret', $global:NinjaOneClientSecret)
                $dict.Add('scope', 'monitoring management')
            }

         
            $formContent = [System.Net.Http.FormUrlEncodedContent]::new($dict)
            $uri = "https://$($global:NinjaOneInstance -replace '/ws','')/ws/oauth/token"

      
            $responseTask = $httpClient.PostAsync($uri, $formContent)
            $response = $responseTask.GetAwaiter().GetResult()

            if (-not $response.IsSuccessStatusCode) {
                Throw "Failed to fetch token: $($response.StatusCode) - $($response.ReasonPhrase)"
            }

            $tokenJson = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $token = $tokenJson | ConvertFrom-Json

            $global:NinjaTokenExpiry = (Get-Date).AddSeconds($token.expires_in)
            $global:NinjaToken = $token

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
        [Parameter(Mandatory = $true)]
        [string]$Instance,
        [Parameter(Mandatory = $true)]
        [string]$ClientID,
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,
        $RefreshToken
    )

    # Set global variables
    $global:NinjaOneInstance = $Instance
    $global:NinjaOneClientId = $ClientID
    $global:NinjaOneClientSecret = $ClientSecret
    $Script:NinjaOneRefreshToken = $RefreshToken
    
    try {
        $null = Get-NinjaOneToken -ErrorAction Stop
    }
    catch {
        Throw "Failed to Connect to NinjaOne: $_"
    }
}

function Invoke-NinjaOneRequestAsync {
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
            $Body = $InputObject | ConvertTo-Json -Depth 100
            if (($InputObject | Measure-Object).Count -eq 1) {
                $Body = '[' + $Body + ']'
            }
        }
        else {
            $Body = $InputObject | ConvertTo-Json -Depth 100
        }
    }

    $httpClient.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $token.access_token)
    $httpClient.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("application/json"))

    $uri = "https://$($global:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams) {"?$QueryParams"})"
    $results = @()

    try {
        if ($Method -in @('GET', 'DELETE')) {
            if ($Paginate) {
                $after = 0
                $pageSize = 1000
                do {
                    $paginatedUri = "https://$($global:NinjaOneInstance)/api/v2/$($Path)?pageSize=$pageSize&after=$after$(if ($QueryParams) {"&$QueryParams"})"
                    $responseTask = $httpClient.GetAsync($paginatedUri)
                    $response = $responseTask.GetAwaiter().GetResult()

                    if (-not $response.IsSuccessStatusCode) {
                        Throw "Request failed: $($response.StatusCode) - $($response.ReasonPhrase)"
                    }

                    $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $result = $content | ConvertFrom-Json
                    $results += $result

                    $resultCount = ($result.id | Measure-Object -Maximum)
                    $after = $resultCount.Maximum
                } while ($resultCount.Count -eq $pageSize)
            }
            else {
                $responseTask = $httpClient.GetAsync($uri)
                $response = $responseTask.GetAwaiter().GetResult()

                if (-not $response.IsSuccessStatusCode) {
                    Throw "Request failed: $($response.StatusCode) - $($response.ReasonPhrase)"
                }

                $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                $results = $content | ConvertFrom-Json
            }
        }
        elseif ($Method -in @('PATCH', 'PUT', 'POST')) {
            $content = [System.Net.Http.StringContent]::new($Body, [System.Text.Encoding]::UTF8, "application/json")
            $responseTask = if ($Method -eq 'POST') {
                $httpClient.PostAsync($uri, $content)
            } elseif ($Method -eq 'PUT') {
                $httpClient.PutAsync($uri, $content)
            } else {
                $httpClient.PatchAsync($uri, $content)
            }

            $response = $responseTask.GetAwaiter().GetResult()

            if (-not $response.IsSuccessStatusCode) {
                Throw "Request failed: $($response.StatusCode) - $($response.ReasonPhrase)"
            }

            $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $results = $content | ConvertFrom-Json
        }
        else {
            Throw 'Unknown Method'
        }
    }
    catch {
        Throw "Error Occurred: $_"
    }

    return $results
}

# Create an InitialSessionState to import functions into runspaces
$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$functionNames = @(
    "Get-NinjaOneToken",
    "Connect-NinjaOne",
    "Invoke-NinjaOneRequestAsync",
    "Convert-SizeToReadable",
    "Convert-UnixTimestampToDateTime"
)

foreach ($funcName in $functionNames) {
    $functionDefinition = (Get-Command $funcName -CommandType Function).Definition
    $initialSessionState.Commands.Add(
        (New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $funcName, $functionDefinition)
    )
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="NinjaOne Software Inventory Exporter" Height="820" Width="1050" WindowStartupLocation="CenterScreen" Background="#1E1E1E">
    <Window.Resources>
        <!-- Style with custom ControlTemplate for TabItem -->
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#2D2D2D"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="6,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1,1,1,0" Margin="2,0">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" 
                                             HorizontalAlignment="Center" ContentSource="Header" Margin="10,2"
                                             TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#2D2D2D"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#3C3C3C"/>
                                <Setter Property="Foreground" Value="Black"/>
                                <Setter TargetName="Border" Property="BorderThickness" Value="1,1,1,2"/>
                                <Setter TargetName="Border" Property="BorderBrush" Value="#555555"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Style with custom ControlTemplate for Button -->
        <Style TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="6,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" Padding="{TemplateBinding Padding}">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" 
                                             HorizontalAlignment="Center" Content="{TemplateBinding Content}"
                                             TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#3C3C3C"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#555555"/>
                                <!-- Keep Foreground White -->
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#3C3C3C"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#2D2D2D"/>
                                <Setter Property="Foreground" Value="Gray"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Style with custom ControlTemplate for ComboBox -->
        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="6,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <Border x:Name="Border" Background="{TemplateBinding Background}" 
                                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
                                <Grid>
                                    <ToggleButton x:Name="ToggleButton" Background="Transparent" 
                                                  BorderBrush="Transparent" BorderThickness="0"
                                                  IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" 
                                                  ClickMode="Press" Focusable="False">
                                        <ToggleButton.ContentTemplate>
                                            <DataTemplate>
                                                <TextBlock Text="{Binding}" 
                                                           Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType={x:Type ComboBox}}}" 
                                                           Margin="3"/>
                                            </DataTemplate>
                                        </ToggleButton.ContentTemplate>
                                    </ToggleButton>
                                    <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" 
                                                     Content="{TemplateBinding SelectionBoxItem}" 
                                                     ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" 
                                                     ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" 
                                                     Margin="3,3,23,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                    <Canvas x:Name="PART_DropDownGlyph" Width="10" Height="5" HorizontalAlignment="Right" 
                                            VerticalAlignment="Center" Margin="0,0,5,0">
                                        <Path x:Name="Path" Width="10" Height="5" Data="M 0 0 L 5 5 L 10 0 Z" Fill="White"/>
                                    </Canvas>
                                </Grid>
                            </Border>
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" 
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Border x:Name="DropDownBorder" Background="#3C3C3C" BorderBrush="#555555" BorderThickness="1">
                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#3C3C3C"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#555555"/>
                                <!-- Keep Foreground White -->
                            </Trigger>
                            <Trigger Property="IsDropDownOpen" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#3C3C3C"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter TargetName="Path" Property="Data" Value="M 0 5 L 5 0 L 10 5 Z"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#2D2D2D"/>
                                <Setter Property="Foreground" Value="Gray"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Style for ComboBoxItem -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="6,2"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#555555"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#555555"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <!-- Style with custom ControlTemplate for CheckBox -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="6,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="16"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Border x:Name="CheckBorder" Width="14" Height="14" Background="{TemplateBinding Background}" 
                                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" VerticalAlignment="Center">
                                <Path x:Name="CheckMark" Width="10" Height="10" Data="M 0 5 L 4 9 L 10 3" Stroke="Transparent" 
                                      StrokeThickness="2" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                            </Border>
                            <ContentPresenter x:Name="ContentSite" Grid.Column="1" Margin="4,0,0,0" 
                                             VerticalAlignment="Center" HorizontalAlignment="Left" 
                                             TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="False">
                                <Setter TargetName="CheckBorder" Property="Background" Value="#3C3C3C"/>
                                <Setter TargetName="CheckMark" Property="Stroke" Value="Transparent"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckBorder" Property="Background" Value="#555555"/>
                                <Setter TargetName="CheckMark" Property="Stroke" Value="White"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckBorder" Property="Background" Value="#555555"/>
                                <!-- Keep Foreground White -->
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="CheckBorder" Property="Background" Value="#2D2D2D"/>
                                <Setter Property="Foreground" Value="Gray"/>
                                <Setter TargetName="CheckMark" Property="Stroke" Value="Gray"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Style for PasswordBox -->
        <Style TargetType="PasswordBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="6,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="PasswordBox">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" Padding="{TemplateBinding Padding}">
                            <ScrollViewer x:Name="PART_ContentHost" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#3C3C3C"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#555555"/>
                                <!-- Keep Foreground White -->
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#2D2D2D"/>
                                <Setter Property="Foreground" Value="Gray"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="300"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar for Exclude List -->
        <Border Grid.Column="0" Background="#2D2D2D" Margin="0,0,15,0" BorderBrush="#444444" BorderThickness="1" Padding="10">
            <StackPanel>
                <Label Content="Exclude Software List" Foreground="White" FontSize="16" Margin="0,0,0,10" ToolTip="List of software to exclude from exports"/>
                <ListBox x:Name="ExcludeListBox" Height="500" Background="#2D2D2D" Foreground="White" BorderBrush="#444444" Margin="0,0,0,15" ToolTip="Select software to remove or add new exclusions"/>
                <Button x:Name="AddExcludeButton" Content="Add" Width="120" Height="30" Margin="0,0,0,10" ToolTip="Add a new software name to the exclude list"/>
                <Button x:Name="RemoveExcludeButton" Content="Remove" Width="120" Height="30" Margin="0,0,0,10" ToolTip="Remove the selected software from the exclude list"/>
                <Button x:Name="RemoveAllExcludeButton" Content="Remove All" Width="120" Height="30" Margin="0,0,0,10" ToolTip="Clear all software from the exclude list"/>
                <Button x:Name="ImportExcludeButton" Content="Import CSV" Width="120" Height="30" Margin="0,0,0,10" ToolTip="Import a CSV file with software names to exclude"/>
            </StackPanel>
        </Border>

        <!-- Tab Control for Main Content -->
        <TabControl Grid.Column="1" Background="#1E1E1E" BorderBrush="#444444">
            <TabItem Header="Software Export">
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
                        <RowDefinition Height="Auto"/> <!-- Added for ProgressBar and Buttons -->
                    </Grid.RowDefinitions>

                    <Label Grid.Row="0" Content="NinjaOne Instance:" Foreground="White" FontSize="14" Margin="0,0,0,10" ToolTip="Select the NinjaOne instance to connect to"/>
                    <ComboBox x:Name="InstanceComboBox" Grid.Row="0" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" SelectedIndex="0" ToolTip="Choose your NinjaOne instance URL">
                        <ComboBoxItem Content="app.ninjarmm.com"/>
                        <ComboBoxItem Content="us2.ninjarmm.com"/>
                        <ComboBoxItem Content="eu.ninjarmm.com"/>
                        <ComboBoxItem Content="ca.ninjarmm.com"/>
                        <ComboBoxItem Content="oc.ninjarmm.com"/>
                    </ComboBox>

                    <Label Grid.Row="1" Content="Client ID:" Foreground="White" FontSize="14" Margin="0,0,0,10" ToolTip="Enter your NinjaOne API Client ID"/>
                    <TextBox x:Name="ClientIdTextBox" Grid.Row="1" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" ToolTip="Input the Client ID for API authentication"/>

                    <Label Grid.Row="2" Content="Client Secret:" Foreground="White" FontSize="14" Margin="0,0,0,10" ToolTip="Enter your NinjaOne API Client Secret"/>
                    <PasswordBox x:Name="ClientSecretPasswordBox" Grid.Row="2" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" ToolTip="Input the Client Secret for API authentication (displayed as secure characters)"/>

                    <Button x:Name="ConnectButton" Grid.Row="3" Content="Connect to API" Width="120" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" ToolTip="Connect to the NinjaOne API with the provided credentials"/>
                    <Ellipse x:Name="ConnectionIndicator" Grid.Row="3" Width="15" Height="15" Margin="280,0,0,15" HorizontalAlignment="Left" Fill="Red" ToolTip="Red: Disconnected, Green: Connected"/>

                    <Label Grid.Row="4" Content="Output Folder:" Foreground="White" FontSize="14" Margin="0,0,0,10" ToolTip="Specify the folder for CSV export files"/>
                    <TextBox x:Name="OutputFolderTextBox" Grid.Row="4" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Text="$CSVOutputFolder" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" ToolTip="Folder path where CSV files will be saved"/>
                    <Button x:Name="BrowseOutputFolderButton" Grid.Row="4" Content="Browse" Width="80" Height="30" Margin="410,0,0,15" HorizontalAlignment="Left" ToolTip="Browse to select the output folder"/>

                    <Label Grid.Row="5" Content="Organizations:" Foreground="White" FontSize="14" Margin="0,0,0,10" ToolTip="Select organizations to include in the export"/>
                    <TextBox x:Name="OrgSearchBox" Grid.Row="5" Width="250" Height="30" Margin="150,0,0,5" HorizontalAlignment="Left" Text="Search organizations..." Foreground="Gray" Background="#3C3C3C" BorderBrush="#555555" ToolTip="Search for organizations by name"/>
                    <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="410,0,0,5" HorizontalAlignment="Left">
                        <Button x:Name="CheckAllButton" Content="Check All" Width="100" Height="30" Margin="0,0,15,0" ToolTip="Select all organizations"/>
                        <Button x:Name="UncheckAllButton" Content="Uncheck All" Width="100" Height="30" ToolTip="Deselect all organizations"/>
                    </StackPanel>
                    <ListBox x:Name="OrgsListBox" Grid.Row="6" Height="250" Margin="0,0,0,15" SelectionMode="Multiple" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" ToolTip="List of organizations; select to include in export">
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
                        <CheckBox x:Name="ExcludeOfficeCheckBox" Grid.Column="0" Content="Exclude Office Apps" Margin="0,0,20,0" ToolTip="Exclude Microsoft Office applications from the export"/>
                        <CheckBox x:Name="ExcludeDotNetCheckBox" Grid.Column="1" Content="Exclude .NET" Margin="0,0,20,0" ToolTip="Exclude .NET Framework and Core from the export"/>
                        <CheckBox x:Name="ExcludeCppCheckBox" Grid.Column="2" Content="Exclude C++ Redist" Margin="0,0,20,0" ToolTip="Exclude Visual C++ Redistributables from the export"/>
                        <Label Grid.Column="3" Content="Version Display:" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0" ToolTip="Choose how software versions are displayed in the export"/>
                        <ComboBox x:Name="VersionDisplayComboBox" Grid.Column="4" Width="200" Height="30" SelectedIndex="0" ToolTip="Select version display option: hide, combine, or list separately">
                            <ComboBoxItem Content="Hide Versions"/>
                            <ComboBoxItem Content="Combine Versions"/>
                            <ComboBoxItem Content="List Each Version"/>
                        </ComboBox>
                    </Grid>

                    <Label Grid.Row="8" Content="Log:" Foreground="White" FontSize="14" Margin="0,0,0,5" ToolTip="View log messages for export progress and errors"/>
                    <TextBox x:Name="LogTextBox" Grid.Row="8" Height="250" Margin="0,0,0,15" AcceptsReturn="True" VerticalScrollBarVisibility="Visible" IsReadOnly="True" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" ToolTip="Log of export activities and status"/>

                    <ProgressBar x:Name="ExportProgressBar" Grid.Row="9" Height="20" Margin="0,0,0,15" Minimum="0" Maximum="100" Value="0" ToolTip="Shows progress of the export operation"/>

                    <StackPanel Grid.Row="10" Orientation="Horizontal">
                        <Button x:Name="RunButton" Content="Run Export" Width="120" Height="30" Margin="0,0,15,0" HorizontalAlignment="Left" IsEnabled="False" ToolTip="Export software inventory to CSV files for selected organizations"/>
                        <Button x:Name="SaveConfigButton" Content="Save Config" Width="120" Height="30" Margin="0,0,0,0" HorizontalAlignment="Left" ToolTip="Save current settings to a configuration file"/>
                    </StackPanel>
                </Grid>
            </TabItem>

            <TabItem Header="All Installed Software">
                <Grid Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Label Grid.Row="0" Content="Search Software:" Foreground="White" FontSize="14" Margin="0,0,0,10" ToolTip="Search for installed software by name"/>
                    <TextBox x:Name="SoftwareSearchBox" Grid.Row="0" Width="250" Height="30" Margin="150,0,0,15" HorizontalAlignment="Left" Text="Search software..." Foreground="Gray" Background="#3C3C3C" BorderBrush="#555555" ToolTip="Enter software name to filter the list"/>

                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="SoftwareCheckAllButton" Content="Check All" Width="100" Height="30" Margin="0,0,15,0" ToolTip="Select all software items"/>
                        <Button x:Name="SoftwareUncheckAllButton" Content="Uncheck All" Width="100" Height="30" ToolTip="Deselect all software items"/>
                    </StackPanel>

                    <ListBox x:Name="SoftwareListBox" Grid.Row="2" Height="500" SelectionMode="Multiple" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" Margin="0,0,0,15" ToolTip="List of installed software; double-click for details, select to export">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <CheckBox Content="{Binding DisplayName}" IsChecked="{Binding IsSelected, RelativeSource={RelativeSource AncestorType={x:Type ListBoxItem}}}" Foreground="White"/>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>

                    <Label Grid.Row="3" Content="Log:" Foreground="White" FontSize="14" Margin="0,0,0,5" ToolTip="View log messages for software selection and export"/>
                    <TextBox x:Name="SoftwareLogTextBox" Grid.Row="3" Height="150" Margin="0,0,0,15" AcceptsReturn="True" VerticalScrollBarVisibility="Visible" IsReadOnly="True" Background="#3C3C3C" Foreground="White" BorderBrush="#555555" ToolTip="Log of software-related activities"/>

                    <Button x:Name="ExportSoftwareButton" Grid.Row="4" Content="Export Selected" Width="120" Height="30" Margin="0,0,0,0" HorizontalAlignment="Left" IsEnabled="False" ToolTip="Export selected software details to a CSV file"/>
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
$clientSecretPasswordBox = $window.FindName("ClientSecretPasswordBox")
$connectButton = $window.FindName("ConnectButton")
$connectionIndicator = $window.FindName("ConnectionIndicator")
$outputFolderTextBox = $window.FindName("OutputFolderTextBox")
$browseOutputFolderButton = $window.FindName("BrowseOutputFolderButton")
$orgSearchBox = $window.FindName("OrgSearchBox")
$checkAllButton = $window.FindName("CheckAllButton")
$uncheckAllButton = $window.FindName("UncheckAllButton")
$orgsListBox = $window.FindName("OrgsListBox")
$logTextBox = $window.FindName("LogTextBox")
$exportProgressBar = $window.FindName("ExportProgressBar")
$runButton = $window.FindName("RunButton")
$saveConfigButton = $window.FindName("SaveConfigButton")
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

if (-not $window -or -not $instanceComboBox -or -not $clientIdTextBox -or -not $clientSecretPasswordBox -or -not $connectButton -or -not $connectionIndicator -or -not $outputFolderTextBox -or -not $browseOutputFolderButton -or -not $orgSearchBox -or -not $checkAllButton -or -not $uncheckAllButton -or -not $orgsListBox -or -not $logTextBox -or -not $exportProgressBar -or -not $runButton -or -not $saveConfigButton -or -not $excludeListBox -or -not $addExcludeButton -or -not $removeExcludeButton -or -not $removeAllExcludeButton -or -not $importExcludeButton -or -not $excludeOfficeCheckBox -or -not $excludeDotNetCheckBox -or -not $excludeCppCheckBox -or -not $versionDisplayComboBox -or -not $softwareSearchBox -or -not $softwareListBox -or -not $softwareCheckAllButton -or -not $softwareUncheckAllButton -or -not $softwareLogTextBox -or -not $exportSoftwareButton) {
    Write-Error "One or more UI elements could not be found. Check XAML naming and structure."
    exit 1
}

# Load configuration if exists
$configFile = Join-Path $CSVOutputFolder "NinjaOneConfig.json"
if (Test-Path $configFile) {
    try {
        $config = Get-Content $configFile | ConvertFrom-Json
        $instanceComboBox.SelectedItem = $instanceComboBox.Items | Where-Object { $_.Content -eq $config.Instance }
        $clientIdTextBox.Text = $config.ClientId
        $outputFolderTextBox.Text = $config.OutputFolder
        $excludeListBox.Items.Clear()
        $config.ExcludeList | ForEach-Object { $excludeListBox.Items.Add($_) }
        $logTextBox.Text += "Loaded configuration from $configFile`r`n"
    }
    catch {
        $logTextBox.Text += "Failed to load configuration: $($_)`r`n"
    }
}
else {
    foreach ($item in $SoftwareWhitelist) {
        $excludeListBox.Items.Add($item) | Out-Null
    }
}

$connectButton.Add_Click({
    $logTextBox.Text = ""
    $connectButton.IsEnabled = $false

    # Create a runspace with the initial session state
    $runspace = [RunspaceFactory]::CreateRunspace($initialSessionState)
    $runspace.Open()
    $powershell = [PowerShell]::Create()
    $powershell.Runspace = $runspace

    $scriptBlock = {
        param($instance, $clientId, $clientSecret, $logTextBox, $connectionIndicator, $orgsListBox, $softwareLogTextBox, $softwareListBox)

        # Re-initialize HttpClient in the runspace
        $httpClientHandler = New-Object System.Net.Http.HttpClientHandler
        $httpClient = New-Object System.Net.Http.HttpClient($httpClientHandler)
        $httpClient.Timeout = [System.TimeSpan]::FromSeconds(30)

        try {
            Connect-NinjaOne -Instance $instance -ClientID $clientId -ClientSecret $clientSecret
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Connected to NinjaOne API`r`n" })
            $connectionIndicator.Dispatcher.Invoke({ $connectionIndicator.Fill = "Green" })

            $AllOrgs = Invoke-NinjaOneRequestAsync -Method Get -Path "organizations"
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Retrieved organizations`r`n" })
            $orgsListBox.Dispatcher.Invoke({
                $orgsListBox.Items.Clear()
                foreach ($org in $AllOrgs) {
                    $orgsListBox.Items.Add([PSCustomObject]@{Name = $org.name})
                }
            })

            $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text = "Retrieving software list...`r`n" })
            $AllSoftware = Invoke-NinjaOneRequestAsync -Method Get -Path "queries/software"
            $cachedSoftware = $AllSoftware.results | Sort-Object name | Group-Object name | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName = $_.Name
                    FullData    = $_.Group
                }
            }
            $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text += "Retrieved $($cachedSoftware.Count) unique software items`r`n" })
            $softwareListBox.Dispatcher.Invoke({
                $softwareListBox.Items.Clear()
                foreach ($software in $cachedSoftware) {
                    $softwareListBox.Items.Add($software)
                }
            })

            return $cachedSoftware, $true, $true
        }
        catch {
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Failed to connect to NinjaOne API: $($_)`r`n" })
            $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text += "Failed to connect to NinjaOne API: $($_)`r`n" })
            $connectionIndicator.Dispatcher.Invoke({ $connectionIndicator.Fill = "Red" })
            return $null, $false, $false
        }
    }

    $powershell.AddScript($scriptBlock)
    $powershell.AddParameter("instance", $instanceComboBox.SelectedItem.Content)
    $powershell.AddParameter("clientId", $clientIdTextBox.Text)
    $powershell.AddParameter("clientSecret", $clientSecretPasswordBox.Password)
    $powershell.AddParameter("logTextBox", $logTextBox)
    $powershell.AddParameter("connectionIndicator", $connectionIndicator)
    $powershell.AddParameter("orgsListBox", $orgsListBox)
    $powershell.AddParameter("softwareLogTextBox", $softwareLogTextBox)
    $powershell.AddParameter("softwareListBox", $softwareListBox)

    $handle = $powershell.BeginInvoke()

    $progressChars = @("|", "/", "-", "\")
    $progressIndex = 0
    while (-not $handle.IsCompleted) {
        $logTextBox.Text = "Connecting to NinjaOne API... $($progressChars[$progressIndex % 4])`r`n"
        $progressIndex++
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
    $logTextBox.Text = "Connecting to NinjaOne API... Done!`r`n" + $logTextBox.Text

    $cachedSoftware, $runEnabled, $exportEnabled = $powershell.EndInvoke($handle)

    $powershell.Dispose()
    $runspace.Close()
    $runspace.Dispose()

    if ($cachedSoftware) {
        $global:CachedSoftware = $cachedSoftware
    }
    $runButton.IsEnabled = $runEnabled
    $exportSoftwareButton.IsEnabled = $exportEnabled
    $connectButton.IsEnabled = $true
})

$browseOutputFolderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Output Folder for CSV Files"
    $folderBrowser.SelectedPath = $outputFolderTextBox.Text
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $outputFolderTextBox.Text = $folderBrowser.SelectedPath
        $logTextBox.Text += "Selected output folder: $($folderBrowser.SelectedPath)`r`n"
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
        $runButton.IsEnabled = $false

        $searchText = $orgSearchBox.Text.ToLower()
        $instance = $instanceComboBox.SelectedItem.Content
        $clientId = $clientIdTextBox.Text
        $clientSecret = $clientSecretPasswordBox.Password

        $runspace = [RunspaceFactory]::CreateRunspace($initialSessionState)
        $runspace.Open()
        $powershell = [PowerShell]::Create()
        $powershell.Runspace = $runspace

        $scriptBlock = {
            param($searchText, $orgsListBox, $instance, $clientId, $clientSecret)

            $httpClientHandler = New-Object System.Net.Http.HttpClientHandler
            $httpClient = New-Object System.Net.Http.HttpClient($httpClientHandler)
            $httpClient.Timeout = [System.TimeSpan]::FromSeconds(30)

            try {
                Connect-NinjaOne -Instance $instance -ClientID $clientId -ClientSecret $clientSecret
                $AllOrgs = Invoke-NinjaOneRequestAsync -Method Get -Path "organizations"
                $filteredOrgs = $AllOrgs | Where-Object { $_.name.ToLower() -like "*$searchText*" }
                $orgsListBox.Dispatcher.Invoke({
                    $orgsListBox.Items.Clear()
                    foreach ($org in $filteredOrgs) {
                        $orgsListBox.Items.Add([PSCustomObject]@{Name = $org.name})
                    }
                })
                return $true
            }
            catch {
                return $false
            }
        }

        $powershell.AddScript($scriptBlock)
        $powershell.AddParameter("searchText", $searchText)
        $powershell.AddParameter("orgsListBox", $orgsListBox)
        $powershell.AddParameter("instance", $instance)
        $powershell.AddParameter("clientId", $clientId)
        $powershell.AddParameter("clientSecret", $clientSecret)

        $handle = $powershell.BeginInvoke()

        $progressChars = @("|", "/", "-", "\")
        $progressIndex = 0
        while (-not $handle.IsCompleted) {
            $logTextBox.Text = "Searching organizations... $($progressChars[$progressIndex % 4])`r`n"
            $progressIndex++
            Start-Sleep -Milliseconds 200
            [System.Windows.Forms.Application]::DoEvents()
        }
        $logTextBox.Text = "Searching organizations... Done!`r`n" + $logTextBox.Text

        $runEnabled = $powershell.EndInvoke($handle)
        $runButton.IsEnabled = $runEnabled

        $powershell.Dispose()
        $runspace.Close()
        $runspace.Dispose()
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
        $logTextBox.Text += "Added '$input' to exclude list`r`n"
    }
})

$removeExcludeButton.Add_Click({
    if ($excludeListBox.SelectedItem) {
        $removedItem = $excludeListBox.SelectedItem
        $excludeListBox.Items.Remove($excludeListBox.SelectedItem)
        $logTextBox.Text += "Removed '$removedItem' from exclude list`r`n"
    }
})

$removeAllExcludeButton.Add_Click({
    $excludeListBox.Items.Clear()
    $logTextBox.Text += "Cleared all items from exclude list`r`n"
})

$importExcludeButton.Add_Click({
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

$saveConfigButton.Add_Click({
    $config = @{
        Instance = $instanceComboBox.SelectedItem.Content
        ClientId = $clientIdTextBox.Text
        OutputFolder = $outputFolderTextBox.Text
        ExcludeList = @($excludeListBox.Items)
    }
    try {
        $config | ConvertTo-Json | Set-Content $configFile
        $logTextBox.Text += "Configuration saved to $configFile`r`n"
    }
    catch {
        $logTextBox.Text += "Failed to save configuration: $($_)`r`n"
    }
})

$runButton.Add_Click({
    $logTextBox.Text = ""
    $exportProgressBar.Value = 0
    $runButton.IsEnabled = $false

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

    # Capture credentials from the UI
    $instance = $instanceComboBox.SelectedItem.Content
    $clientId = $clientIdTextBox.Text
    $clientSecret = $clientSecretPasswordBox.Password

    $runspace = [RunspaceFactory]::CreateRunspace($initialSessionState)
    $runspace.Open()
    $powershell = [PowerShell]::Create()
    $powershell.Runspace = $runspace

    $scriptBlock = {
        param($orgsListBox, $excludeList, $versionDisplayComboBox, $CSVOutputFolder, $logTextBox, $exportProgressBar, $instance, $clientId, $clientSecret)

        $httpClientHandler = New-Object System.Net.Http.HttpClientHandler
        $httpClient = New-Object System.Net.Http.HttpClient($httpClientHandler)
        $httpClient.Timeout = [System.TimeSpan]::FromSeconds(30)

        # Call Connect-NinjaOne to set global variables in this runspace
        try {
            Connect-NinjaOne -Instance $instance -ClientID $clientId -ClientSecret $clientSecret
        }
        catch {
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Failed to connect to NinjaOne API: $($_)`r`n" })
            return $false
        }

        $MasterList = @()
        $orgList = if ($orgsListBox.SelectedItems.Count -eq 0) {
            $AllOrgs = Invoke-NinjaOneRequestAsync -Method Get -Path "organizations"
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Processing all organizations`r`n" })
            $AllOrgs
        }
        else {
            $AllOrgs = Invoke-NinjaOneRequestAsync -Method Get -Path "organizations"
            $selectedOrgs = $AllOrgs | Where-Object { $orgsListBox.SelectedItems.Name -contains $_.name }
            if (-not $selectedOrgs) {
                $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "No matching organizations selected`r`n" })
                return $null
            }
            $selectedOrgs
        }

        $totalOrgs = $orgList.Count
        $processedOrgs = 0

        foreach ($Org in $orgList) {
            $OrgID = $Org.id
            $OrgName = $Org.name
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Processing $($OrgName)`r`n" })

            try {
                $SoftwarePath = "queries/software?df=org=$($OrgID)"
                $AllSoftwareForOrg = Invoke-NinjaOneRequestAsync -Method Get -Path $SoftwarePath
            }
            catch {
                $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Failed to retrieve software for $($OrgName): $($_)`r`n" })
                $processedOrgs++
                $exportProgressBar.Dispatcher.Invoke({ $exportProgressBar.Value = ($processedOrgs / $totalOrgs) * 100 })
                continue
            }

            if (-not $AllSoftwareForOrg -or -not $AllSoftwareForOrg.results) {
                $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "No software results for $($OrgName)`r`n" })
                $processedOrgs++
                $exportProgressBar.Dispatcher.Invoke({ $exportProgressBar.Value = ($processedOrgs / $totalOrgs) * 100 })
                continue
            }

            try {
                $DevicesPath = "devices?df=org=$($OrgID)"
                $AllDevicesForOrg = Invoke-NinjaOneRequestAsync -Method Get -Path $DevicesPath
            }
            catch {
                $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Failed to retrieve devices for $($OrgName): $($_)`r`n" })
                $processedOrgs++
                $exportProgressBar.Dispatcher.Invoke({ $exportProgressBar.Value = ($processedOrgs / $totalOrgs) * 100 })
                continue
            }

            $deviceLookup = @{}
            foreach ($device in $AllDevicesForOrg) {
                $deviceLookup[$device.id] = $device.systemName
            }

            $filteredSoftware = $AllSoftwareForOrg.results | Where-Object { $excludeList -notcontains $_.name }

            $versionDisplayIndex = $versionDisplayComboBox.Dispatcher.Invoke({ $versionDisplayComboBox.SelectedIndex })
            switch ($versionDisplayIndex) {
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
                            installDate  = if ($_.Group[0].installDate) { ([DateTime]::Parse($_.Group[0].installDate)).ToString("MM/dd/yyyy") } else { "" }
                            location     = $_.Group[0].location
                            size         = if ($_.Group[0].size) { Convert-SizeToReadable -Bytes $_.Group[0].size } else { "" }
                            lastUpdated  = if ($_.Group[0].timestamp) { Convert-UnixTimestampToDateTime -UnixTimestamp $_.Group[0].timestamp } else { "" }
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
                            installDate  = if ($_.Group[0].installDate) { ([DateTime]::Parse($_.Group[0].installDate)).ToString("MM/dd/yyyy") } else { "" }
                            location     = $_.Group[0].location
                            size         = if ($_.Group[0].size) { Convert-SizeToReadable -Bytes $_.Group[0].size } else { "" }
                            lastUpdated  = if ($_.Group[0].timestamp) { Convert-UnixTimestampToDateTime -UnixTimestamp $_.Group[0].timestamp } else { "" }
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
                            installDate  = if ($_.installDate) { ([DateTime]::Parse($_.installDate)).ToString("MM/dd/yyyy") } else { "" }
                            location     = $_.location
                            size         = if ($_.size) { Convert-SizeToReadable -Bytes $_.size } else { "" }
                            lastUpdated  = if ($_.timestamp) { Convert-UnixTimestampToDateTime -UnixTimestamp $_.timestamp } else { "" }
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
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "CSV exported for $($OrgName): $($filepath)`r`n" })

            $processedOrgs++
            $exportProgressBar.Dispatcher.Invoke({ $exportProgressBar.Value = ($processedOrgs / $totalOrgs) * 100 })
        }

        if ($orgsListBox.SelectedItems.Count -eq 0 -and $MasterList) {
            $masterFilename = "Software Inventory Export - $($timestamp) - AllOrgs.csv"
            $masterFilepath = Join-Path -Path $CSVOutputFolder -ChildPath $masterFilename
            $MasterList | Export-Csv -Path $masterFilepath -NoTypeInformation
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "Master list exported: $($masterFilepath)`r`n" })
        }
        elseif (-not $MasterList) {
            $logTextBox.Dispatcher.Invoke({ $logTextBox.Text += "No data collected to export.`r`n" })
        }

        return $true
    }

    $powershell.AddScript($scriptBlock)
    $powershell.AddParameter("orgsListBox", $orgsListBox)
    $powershell.AddParameter("excludeList", $excludeList)
    $powershell.AddParameter("versionDisplayComboBox", $versionDisplayComboBox)
    $powershell.AddParameter("CSVOutputFolder", $CSVOutputFolder)
    $powershell.AddParameter("logTextBox", $logTextBox)
    $powershell.AddParameter("exportProgressBar", $exportProgressBar)
    $powershell.AddParameter("instance", $instance)
    $powershell.AddParameter("clientId", $clientId)
    $powershell.AddParameter("clientSecret", $clientSecret)

    $handle = $powershell.BeginInvoke()

    $progressChars = @("|", "/", "-", "\")
    $progressIndex = 0
    while (-not $handle.IsCompleted) {
        $logTextBox.Text = "Exporting software inventory... $($progressChars[$progressIndex % 4])`r`n"
        $progressIndex++
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
    $logTextBox.Text = "Exporting software inventory... Done!`r`n" + $logTextBox.Text

    $success = $powershell.EndInvoke($handle)
    $runButton.IsEnabled = $success

    $powershell.Dispose()
    $runspace.Close()
    $runspace.Dispose()
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

$softwareListBox.Add_MouseDoubleClick({
    if ($softwareListBox.SelectedItem) {
        $software = $softwareListBox.SelectedItem
        $detailWindowXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Details for $($software.DisplayName)" Width="900" Height="400" Background="#1E1E1E" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#3D3D3D"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#3C3C3C"/>
        </Style>
    </Window.Resources>
    <DataGrid x:Name="SoftwareDetailGrid" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True"
              Background="#3C3C3C" Foreground="White" BorderBrush="#555555"
              RowBackground="#3C3C3C" AlternatingRowBackground="#2D2D2D">
        <DataGrid.Columns>
            <DataGridTextColumn Header="Device" Binding="{Binding Device}" Width="120"/>
            <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="80"/>
            <DataGridTextColumn Header="Publisher" Binding="{Binding Publisher}" Width="120"/>
            <DataGridTextColumn Header="Product Code" Binding="{Binding ProductCode}" Width="120"/>
            <DataGridTextColumn Header="Install Date" Binding="{Binding InstallDate}" Width="90"/>
            <DataGridTextColumn Header="Location" Binding="{Binding Location}" Width="150"/>
            <DataGridTextColumn Header="Size" Binding="{Binding Size}" Width="80"/>
            <DataGridTextColumn Header="Last Updated" Binding="{Binding LastUpdated}" Width="120"/>
        </DataGrid.Columns>
    </DataGrid>
</Window>
"@

        try {
            $detailWindow = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$detailWindowXaml)))
        }
        catch {
            $softwareLogTextBox.Text += "Failed to load detail window: $($_)`r`n"
            return
        }

        $grid = $detailWindow.FindName("SoftwareDetailGrid")

        # Fetch device data asynchronously
        $instance = $instanceComboBox.SelectedItem.Content
        $clientId = $clientIdTextBox.Text
        $clientSecret = $clientSecretPasswordBox.Password

        $runspace = [RunspaceFactory]::CreateRunspace($initialSessionState)
        $runspace.Open()
        $powershell = [PowerShell]::Create()
        $powershell.Runspace = $runspace

        $scriptBlock = {
            param($software, $softwareLogTextBox, $instance, $clientId, $clientSecret)

            $httpClientHandler = New-Object System.Net.Http.HttpClientHandler
            $httpClient = New-Object System.Net.Http.HttpClient($httpClientHandler)
            $httpClient.Timeout = [System.TimeSpan]::FromSeconds(30)

            try {
                Connect-NinjaOne -Instance $instance -ClientID $clientId -ClientSecret $clientSecret
                $AllDevices = Invoke-NinjaOneRequestAsync -Method Get -Path "devices"
                $deviceLookup = @{}
                foreach ($device in $AllDevices) {
                    $deviceLookup[$device.id] = $device.systemName
                }

                $dataItems = @(@($software.FullData) | ForEach-Object {
                    [PSCustomObject]@{
                        Device = $deviceLookup[$_.deviceId] ?? "Unknown Device"
                        Version = $_.version
                        Publisher = $_.publisher
                        ProductCode = $_.productCode
                        InstallDate = if ($_.installDate) { ([DateTime]::Parse($_.installDate)).ToString("MM/dd/yyyy") } else { "" }
                        Location = $_.location
                        Size = if ($_.size) { Convert-SizeToReadable -Bytes $_.size } else { "" }
                        LastUpdated = if ($_.timestamp) { Convert-UnixTimestampToDateTime -UnixTimestamp $_.timestamp } else { "" }
                    }
                })

                return $dataItems
            }
            catch {
                $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text += "Failed to fetch device data: $($_)`r`n" })
                return $null
            }
        }

        $powershell.AddScript($scriptBlock)
        $powershell.AddParameter("software", $software)
        $powershell.AddParameter("softwareLogTextBox", $softwareLogTextBox)
        $powershell.AddParameter("instance", $instance)
        $powershell.AddParameter("clientId", $clientId)
        $powershell.AddParameter("clientSecret", $clientSecret)

        $handle = $powershell.BeginInvoke()

        $progressChars = @("|", "/", "-", "\")
        $progressIndex = 0
        while (-not $handle.IsCompleted) {
            $softwareLogTextBox.Text = "Loading device details... $($progressChars[$progressIndex % 4])`r`n"
            $progressIndex++
            Start-Sleep -Milliseconds 200
            [System.Windows.Forms.Application]::DoEvents()
        }
        $softwareLogTextBox.Text = "Loading device details... Done!`r`n" + $softwareLogTextBox.Text

        $dataItems = $powershell.EndInvoke($handle)

        $powershell.Dispose()
        $runspace.Close()
        $runspace.Dispose()

        if ($dataItems) {
            $grid.ItemsSource = $dataItems
            $detailWindow.ShowDialog() | Out-Null
        }
    }
})

$exportSoftwareButton.Add_Click({
    $softwareLogTextBox.Text = ""
    $exportSoftwareButton.IsEnabled = $false

    $CSVOutputFolder = $outputFolderTextBox.Text
    $selectedSoftware = $softwareListBox.SelectedItems

    if ($selectedSoftware.Count -eq 0) {
        $softwareLogTextBox.Text += "No software selected for export`r`n"
        $exportSoftwareButton.IsEnabled = $true
        return
    }

    $instance = $instanceComboBox.SelectedItem.Content
    $clientId = $clientIdTextBox.Text
    $clientSecret = $clientSecretPasswordBox.Password

    $runspace = [RunspaceFactory]::CreateRunspace($initialSessionState)
    $runspace.Open()
    $powershell = [PowerShell]::Create()
    $powershell.Runspace = $runspace

    $scriptBlock = {
        param($selectedSoftware, $CSVOutputFolder, $softwareLogTextBox, $instance, $clientId, $clientSecret)

        $httpClientHandler = New-Object System.Net.Http.HttpClientHandler
        $httpClient = New-Object System.Net.Http.HttpClient($httpClientHandler)
        $httpClient.Timeout = [System.TimeSpan]::FromSeconds(30)

        try {
            Connect-NinjaOne -Instance $instance -ClientID $clientId -ClientSecret $clientSecret
            $AllDevices = Invoke-NinjaOneRequestAsync -Method Get -Path "devices"
            $deviceLookup = @{}
            foreach ($device in $AllDevices) {
                $deviceLookup[$device.id] = $device.systemName
            }

            $exportData = @()
            foreach ($software in $selectedSoftware) {
                $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text += "Processing $($software.DisplayName)`r`n" })
                foreach ($instance in $software.FullData) {
                    $deviceName = $deviceLookup[$instance.deviceId] ?? "Unknown Device"
                    $exportData += [PSCustomObject]@{
                        SoftwareName = $software.DisplayName
                        Version      = $instance.version
                        Publisher    = $instance.publisher
                        DeviceName   = $deviceName
                        DeviceId     = $instance.deviceId
                        ProductCode  = $instance.productCode
                        InstallDate  = if ($instance.installDate) { ([DateTime]::Parse($instance.installDate)).ToString("MM/dd/yyyy") } else { "" }
                        Location     = $instance.location
                        Size         = if ($instance.size) { Convert-SizeToReadable -Bytes $instance.size } else { "" }
                        LastUpdated  = if ($instance.timestamp) { Convert-UnixTimestampToDateTime -UnixTimestamp $instance.timestamp } else { "" }
                    }
                }
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH-mm"
            $filename = "Selected Software Export - $($timestamp).csv"
            $filepath = Join-Path -Path $CSVOutputFolder -ChildPath $filename
            
            $exportData | Export-Csv -Path $filepath -NoTypeInformation
            $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text += "Exported $($selectedSoftware.Count) software items to: $($filepath)`r`n" })

            return $true
        }
        catch {
            $softwareLogTextBox.Dispatcher.Invoke({ $softwareLogTextBox.Text += "Failed to export software: $($_)`r`n" })
            return $false
        }
    }
    $powershell.AddScript($scriptBlock)
    $powershell.AddParameter("selectedSoftware", $selectedSoftware)
    $powershell.AddParameter("CSVOutputFolder", $CSVOutputFolder)
    $powershell.AddParameter("softwareLogTextBox", $softwareLogTextBox)
    $powershell.AddParameter("instance", $instance)
    $powershell.AddParameter("clientId", $clientId)  
    $powershell.AddParameter("clientSecret", $clientSecret)

    $handle = $powershell.BeginInvoke()

    $progressChars = @("|", "/", "-", "\")
    $progressIndex = 0
    while (-not $handle.IsCompleted) {
        $softwareLogTextBox.Text = "Exporting selected software... $($progressChars[$progressIndex % 4])`r`n"
        $progressIndex++
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
    $softwareLogTextBox.Text = "Exporting selected software... Done!`r`n" + $softwareLogTextBox.Text

    $success = $powershell.EndInvoke($handle)
    $exportSoftwareButton.IsEnabled = $success

    $powershell.Dispose()
    $runspace.Close()
    $runspace.Dispose()
})

$window.ShowDialog() | Out-Null

# Cleanup HttpClient
$httpClient.Dispose()
