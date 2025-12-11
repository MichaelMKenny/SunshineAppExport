# Load assemblies
Add-Type -AssemblyName System.Windows.Forms

function GetMainMenuItems {
    param(
        $getMainMenuItemsArgs
    )

    $menuItem1 = New-Object Playnite.SDK.Plugins.ScriptMainMenuItem
    $menuItem1.Description = "Export selected games"
    $menuItem1.FunctionName = "SunshineExport"
    $menuItem1.MenuSection = "@Sunshine App Export"

    $menuItem2 = New-Object Playnite.SDK.Plugins.ScriptMainMenuItem
    $menuItem2.Description = "Configure Sunshine Export"
    $menuItem2.FunctionName = "Show-ConfigurationDialog"
    $menuItem2.MenuSection = "@Sunshine App Export"
    
    return $menuItem1, $menuItem2
}

function Get-PluginConfig {
    $configPath = Join-Path -Path (Split-Path $script:MyInvocation.MyCommand.Path) -ChildPath "config.json"
    if (Test-Path $configPath) {
        return Get-Content $configPath | ConvertFrom-Json
    }
    return @{
        SunshineUrl             = "https://localhost:47990"
        SunshineUser            = ""
        SunshinePass            = ""
        IgnoreCertificateErrors = $true
        SyncOnStartup           = $false
        KeepUpToDate            = $false
    }
}

function Set-PluginConfig {
    param($Config)
    $configPath = Join-Path -Path (Split-Path $script:MyInvocation.MyCommand.Path) -ChildPath "config.json"
    $Config | ConvertTo-Json | Set-Content $configPath
}

function Show-ConfigurationDialog {
    param($showConfigArgs)

    $config = Get-PluginConfig

    $windowCreationOptions = New-Object Playnite.SDK.WindowCreationOptions
    $windowCreationOptions.ShowMinimizeButton = $false
    $windowCreationOptions.ShowMaximizeButton = $false
    
    $window = $PlayniteApi.Dialogs.CreateWindow($windowCreationOptions)
    $window.Title = "Sunshine Configuration"
    $window.SizeToContent = "WidthAndHeight"
    $window.ResizeMode = "NoResize"
    
    [xml]$xaml = @"
<UserControl
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <StackPanel Margin="20">
        <TextBlock Text="Sunshine URL:" Margin="0,0,0,5"/>
        <TextBox x:Name="SunshineUrl" Width="300" Margin="0,0,0,15"/>
        
        <TextBlock Text="Username:" Margin="0,0,0,5"/>
        <TextBox x:Name="SunshineUser" Width="300" Margin="0,0,0,15"/>
        
        <TextBlock Text="Password:" Margin="0,0,0,5"/>
        <PasswordBox x:Name="SunshinePass" Width="300" Margin="0,0,0,15"/>
        
        <CheckBox x:Name="IgnoreCertErrors" Content="Ignore Certificate Errors" Margin="0,0,0,15"/>
        <CheckBox x:Name="SyncOnStartup" Content="Sync on Playnite Startup" Margin="0,0,0,15"/>
        <CheckBox x:Name="KeepUpToDate" Content="Keep up to date (Real-time sync)" Margin="0,0,0,15"/>
        
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="SaveButton" Content="Save" Width="75" Margin="0,0,10,0"/>
            <Button x:Name="CancelButton" Content="Cancel" Width="75"/>
        </StackPanel>
    </StackPanel>
</UserControl>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window.Content = [Windows.Markup.XamlReader]::Load($reader)

    # Populate fields
    $window.Content.FindName("SunshineUrl").Text = $config.SunshineUrl
    $window.Content.FindName("SunshineUser").Text = $config.SunshineUser
    $window.Content.FindName("SunshinePass").Password = $config.SunshinePass
    $window.Content.FindName("IgnoreCertErrors").IsChecked = $config.IgnoreCertificateErrors
    $window.Content.FindName("SyncOnStartup").IsChecked = $config.SyncOnStartup
    $window.Content.FindName("KeepUpToDate").IsChecked = $config.KeepUpToDate

    # Event Handlers
    $window.Content.FindName("SaveButton").Add_Click({
            $newConfig = @{
                SunshineUrl             = $window.Content.FindName("SunshineUrl").Text
                SunshineUser            = $window.Content.FindName("SunshineUser").Text
                SunshinePass            = $window.Content.FindName("SunshinePass").Password
                IgnoreCertificateErrors = $window.Content.FindName("IgnoreCertErrors").IsChecked
                SyncOnStartup           = $window.Content.FindName("SyncOnStartup").IsChecked
                KeepUpToDate            = $window.Content.FindName("KeepUpToDate").IsChecked
            }
            Set-PluginConfig -Config $newConfig
            $window.Close()
        })

    $window.Content.FindName("CancelButton").Add_Click({
            $window.Close()
        })

    $window.Owner = $PlayniteApi.Dialogs.GetCurrentAppWindow()
    $window.WindowStartupLocation = "CenterOwner"
    $window.ShowDialog()
}

# Helper type for safe SSL validation without using PowerShell ScriptBlocks
try {
    $code = @"
    using System;
    using System.Net.Http;
    using System.Security.Cryptography.X509Certificates;
    using System.Net.Security;

    namespace SunshineExport {
        public class CertificateBypass {
            public static Func<HttpRequestMessage, X509Certificate2, X509Chain, SslPolicyErrors, bool> Callback = 
                (msg, cert, chain, sslErrors) => true;
        }
    }
"@
    if (-not ([System.Management.Automation.PSTypeName]'SunshineExport.CertificateBypass').Type) {
        Add-Type -TypeDefinition $code -Language CSharp -ReferencedAssemblies System.Net.Http
    }
}
catch {
    # Ignored, type might already exist or assembly issue
    $__logger.Error("Failed to add CertificateBypass type: $_")
}

function Invoke-SunshineRequest {
    param(
        [string]$Method,
        [string]$Endpoint,
        [object]$Body = $null
    )

    $config = Get-PluginConfig
    $baseUri = $config.SunshineUrl.TrimEnd('/')
    $uri = "$baseUri$Endpoint"

    $handler = New-Object System.Net.Http.HttpClientHandler
    if ($config.IgnoreCertificateErrors) {
        try {
            # Use the compiled C# delegate to avoid Runspace issues
            $handler.ServerCertificateCustomValidationCallback = [SunshineExport.CertificateBypass]::Callback
        }
        catch {
            $__logger.Error("Failed to set SSL bypass callback: $_")
        }
    }

    $client = New-Object System.Net.Http.HttpClient($handler)
    try {
        $client.BaseAddress = New-Object System.Uri($baseUri)
        
        # Auth
        if (![string]::IsNullOrEmpty($config.SunshineUser) -or ![string]::IsNullOrEmpty($config.SunshinePass)) {
            $authBytes = [System.Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $config.SunshineUser, $config.SunshinePass))
            $authString = [System.Convert]::ToBase64String($authBytes)
            $client.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Basic", $authString)
        }

        # Content
        $content = $null
        if ($Body) {
            $json = $Body | ConvertTo-Json -Depth 10
            $content = New-Object System.Net.Http.StringContent($json, [System.Text.Encoding]::UTF8, "application/json")
        }

        $__logger.Info("Sunshine DEBUG: Invoking $Method request to $uri")

        # Execute
        if ($Method -eq "GET") {
            $task = $client.GetAsync($Endpoint)
        }
        elseif ($Method -eq "POST") {
            $task = $client.PostAsync($Endpoint, $content)
        }
        elseif ($Method -eq "DELETE") {
            $task = $client.DeleteAsync($Endpoint)
        }
        
        $task.Wait()
        $result = $task.Result
        
        $responseBodyTask = $result.Content.ReadAsStringAsync()
        $responseBodyTask.Wait()
        $responseBody = $responseBodyTask.Result

        if (-not $result.IsSuccessStatusCode) {
            $__logger.Error("Sunshine DEBUG: API Error ($($result.StatusCode)): $responseBody")
            throw "Sunshine API returned $($result.StatusCode): $responseBody"
        }

        $__logger.Info("Sunshine DEBUG: Request successful.")
        
        try {
            if ([string]::IsNullOrWhiteSpace($responseBody)) {
                return $null
            }
            return $responseBody | ConvertFrom-Json
        }
        catch {
            return $responseBody
        }
    }
    finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Get-SunshineApps {
    return Invoke-SunshineRequest -Method "GET" -Endpoint "/api/apps"
}

function Add-SunshineApp {
    param($App)
    return Invoke-SunshineRequest -Method "POST" -Endpoint "/api/apps" -Body $App
}

function Remove-SunshineApp {
    param($Index)
    return Invoke-SunshineRequest -Method "DELETE" -Endpoint "/api/apps/$Index"
}

function OnApplicationStarted {
    $config = Get-PluginConfig
    if (-not $config.SyncOnStartup) {
        return
    }

    $installedGames = $PlayniteApi.Database.Games | Where-Object { $_.IsInstalled }
    # Convert to List for Sync-Games
    $gamesList = New-Object System.Collections.Generic.List[Playnite.SDK.Models.Game]
    if ($installedGames) {
        foreach ($game in $installedGames) {
            $gamesList.Add($game)
        }
    }
    
    $count = Sync-Games -GamesToSync $gamesList -RemoveMissing $true
    $__logger.Info("Sunshine Export: Synced $count games on startup.")
}

function OnGameInstalled {
    param($Event)
    
    $Game = $Event
    # Check if input is the EventArgs wrapper (OnGameInstalledEventArgs)
    if ($Event.PSObject.Properties['Game']) {
        $Game = $Event.Game
    }

    $config = Get-PluginConfig
    if ($config.KeepUpToDate) {
        $gamesList = New-Object System.Collections.Generic.List[Playnite.SDK.Models.Game]
        try {
            $gamesList.Add($Game)
            Sync-Games -GamesToSync $gamesList -RemoveMissing $false
            $__logger.Info("Sunshine Export: Auto-synced installed game: $($Game.Name)")
        }
        catch {
            $__logger.Error("Sunshine Export: Failed to process installed game. Input Type: $($Event.GetType().FullName). Error: $_")
        }
    }
}

function OnGameUninstalled {
    param($Event)
    
    $Game = $Event
    # Check if input is the EventArgs wrapper (OnGameUninstalledEventArgs)
    if ($Event.PSObject.Properties['Game']) {
        $Game = $Event.Game
    }

    $config = Get-PluginConfig
    if ($config.KeepUpToDate) {
        try {
            $sunshineAppsResponse = Get-SunshineApps
            $sunshineApps = $sunshineAppsResponse.apps
        }
        catch {
            $__logger.Error("Sunshine Export: Failed to fetch apps for uninstall: $_")
            return
        }

        # Check if we have apps to check
        if ($null -eq $sunshineApps) {
            $__logger.Info("Sunshine Export: No apps found in Sunshine to check for removal.")
            return
        }

        $indexToRemove = -1
        for ($i = 0; $i -lt $sunshineApps.Count; $i++) {
            $app = $sunshineApps[$i]
            
            # Safely check for detached property and its content
            if ($null -ne $app -and $null -ne $app.detached -and $app.detached.Count -gt 0) {
                # GetGameIdFromCmd is defined later in the file, but visible in module scope
                $detachedCmd = $app.detached[0]
                if ($null -ne $detachedCmd) {
                    $existingGameId = GetGameIdFromCmd($detachedCmd)
                    if ($existingGameId -eq $Game.Id.ToString()) {
                        $indexToRemove = $i
                        break
                    }
                }
            }
        }
        
        if ($indexToRemove -ge 0) {
            try {
                Remove-SunshineApp -Index $indexToRemove
                $__logger.Info("Sunshine Export: Auto-removed uninstalled game: $($Game.Name)")
            }
            catch {
                $__logger.Error("Sunshine Export: Failed to remove uninstalled game: $_")
            }
        }
    }
}

function SunshineExport {
    param(
        $scriptMainMenuItemActionArgs
    )

    $selectedInstalledGames = $PlayniteApi.MainView.SelectedGames | Where-Object { $_.IsInstalled }
    
    $gamesList = New-Object System.Collections.Generic.List[Playnite.SDK.Models.Game]
    if ($selectedInstalledGames) {
        foreach ($game in $selectedInstalledGames) {
            $gamesList.Add($game)
        }
    }

    $shortcutsCreatedCount = Sync-Games -GamesToSync $gamesList -RemoveMissing $false
    
    $totalSelected = 0
    if ($PlayniteApi.MainView.SelectedGames) { $totalSelected = $PlayniteApi.MainView.SelectedGames.Count }

    $message = "Exported $shortcutsCreatedCount games to Sunshine."
    if ($totalSelected -gt $gamesList.Count) {
        $skipped = $totalSelected - $gamesList.Count
        $message += "`n($skipped non-installed games skipped)"
    }

    $PlayniteApi.Dialogs.ShowMessage($message, "Sunshine Export")
}

function GetGameIdFromCmd([string]$cmd) {
    $parts = $cmd -split " --start "
    if ($parts.Count -gt 1) {
        return ($parts[1] -split " ")[0]
    }
    else {
        return ""
    }
}

function Sync-Games {
    param(
        [System.Collections.Generic.List[Playnite.SDK.Models.Game]]$GamesToSync,
        [bool]$RemoveMissing = $false
    )

    # Load assemblies
    Add-Type -AssemblyName System.Drawing
    $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
    
    # Set paths
    $playniteExecutablePath = Join-Path -Path $PlayniteApi.Paths.ApplicationPath -ChildPath "Playnite.DesktopApp.exe"
    $appAssetsPath = Join-Path -Path $env:LocalAppData -ChildPath "Sunshine Playnite App Export\Apps"
    if (!(Test-Path $appAssetsPath -PathType Container)) {
        New-Item -ItemType Container -Path $appAssetsPath -Force
    }

    $shortcutsCreatedCount = 0
    
    try {
        $sunshineAppsResponse = Get-SunshineApps
        $sunshineApps = $sunshineAppsResponse.apps
    }
    catch {
        $PlayniteApi.Dialogs.ShowErrorMessage("Failed to connect to Sunshine. Please check your configuration.", "Sunshine Export Error")
        return 0
    }

    # Create a hashset of Game IDs to sync for O(1) lookup
    $gamesToSyncIds = New-Object System.Collections.Generic.HashSet[string]
    foreach ($game in $GamesToSync) {
        $gamesToSyncIds.Add($game.Id.ToString()) | Out-Null
    }

    # 1. Add/Update Games
    foreach ($game in $GamesToSync) {
        $gameLaunchCmd = "`"$playniteExecutablePath`" --start $($game.Id)"

        # Set cover path and create blank file
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.Id, "box-art.png")
        if (!(Test-Path $sunshineGameCoverPath -PathType Container)) {
            New-Item -ItemType File -Path $sunshineGameCoverPath -Force | Out-Null
        }

        if ($null -ne $game.CoverImage) {
            $sourceCover = $PlayniteApi.Database.GetFullFilePath($game.CoverImage)
            if (($game.CoverImage -notmatch "^http") -and (Test-Path $sourceCover -PathType Leaf)) {
                if ([System.IO.Path]::GetExtension($game.CoverImage) -eq ".png") {
                    Copy-Item $sourceCover $sunshineGameCoverPath -Force
                }
                else {
                    # Convert cover image to compatible PNG image format
                    try {
                        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                        $bitmap.BeginInit()
                        $bitmap.UriSource = New-Object System.Uri($sourceCover, [System.UriKind]::Absolute)
                        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bitmap.EndInit()
                    
                        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
                        $frame = [System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap)
                        $encoder.Frames.Add($frame)
                    
                        $fileStream = New-Object System.IO.FileStream($sunshineGameCoverPath, [System.IO.FileMode]::Create)
                        $encoder.Save($fileStream)
                        $fileStream.Close()
                    }
                    catch {
                        if ($null -ne $bitmap) { $bitmap = $null }
                        if ($null -ne $fileStream) { $fileStream.Close() }
                        $errorMessage = $_.Exception.Message
                        $__logger.Info("Error converting cover image of `"$($game.Name)`". Error: $errorMessage")
                    }
                }
            }
        }

        # Check if app already exists in Sunshine
        $existingAppIndex = -1
        for ($i = 0; $i -lt $sunshineApps.Count; $i++) {
            $app = $sunshineApps[$i]
            if ($app.detached -and $app.detached.Length -gt 0) {
                $existingGameId = GetGameIdFromCmd($app.detached[0])
                if ($existingGameId -eq $game.Id.ToString()) {
                    $existingAppIndex = $i
                    break
                }
            }
        }

        $newApp = @{
            name         = $game.Name
            detached     = @($gameLaunchCmd)
            "image-path" = $sunshineGameCoverPath
            index        = $existingAppIndex
        }

        try {
            Add-SunshineApp -App $newApp
            $shortcutsCreatedCount++
        }
        catch {
            $__logger.Error("Failed to add/update app for game $($game.Name): $_")
        }
    }

    # 2. Remove Missing Games (if requested)
    if ($RemoveMissing) {
       
        for ($i = $sunshineApps.Count - 1; $i -ge 0; $i--) {
            $app = $sunshineApps[$i]
            if ($app.detached -and $app.detached.Length -gt 0) {
                $existingGameId = GetGameIdFromCmd($app.detached[0])
                if (![string]::IsNullOrEmpty($existingGameId)) {
                    # It's a Playnite exported app. Check if it should be synced.
                    if (-not $gamesToSyncIds.Contains($existingGameId)) {
                        try {
                            Remove-SunshineApp -Index $i
                            $__logger.Info("Removed Sunshine app for game ID: $existingGameId")
                        }
                        catch {
                            $__logger.Error("Failed to remove Sunshine app at index ${i}: $_")
                        }
                    }
                }
            }
        }
    }

    return $shortcutsCreatedCount
}
