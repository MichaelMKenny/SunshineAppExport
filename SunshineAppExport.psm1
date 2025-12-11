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

    # Event Handlers
    $window.Content.FindName("SaveButton").Add_Click({
            $newConfig = @{
                SunshineUrl             = $window.Content.FindName("SunshineUrl").Text
                SunshineUser            = $window.Content.FindName("SunshineUser").Text
                SunshinePass            = $window.Content.FindName("SunshinePass").Password
                IgnoreCertificateErrors = $window.Content.FindName("IgnoreCertErrors").IsChecked
                SyncOnStartup           = $window.Content.FindName("SyncOnStartup").IsChecked
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

function Invoke-SunshineRequest {
    param(
        [string]$Method,
        [string]$Endpoint,
        [object]$Body = $null
    )

    $config = Get-PluginConfig
    $baseUri = $config.SunshineUrl.TrimEnd('/')
    $uri = "$baseUri$Endpoint"

    $headers = @{}
    if (![string]::IsNullOrEmpty($config.SunshineUser) -or ![string]::IsNullOrEmpty($config.SunshinePass)) {
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $config.SunshineUser, $config.SunshinePass)))
        $headers.Add("Authorization", "Basic $auth")
    }

    $params = @{
        Uri         = $uri
        Method      = $Method
        Headers     = $headers
        ContentType = "application/json"
    }

    if ($Body) {
        $params.Body = $Body | ConvertTo-Json -Depth 10
    }

    if ($config.IgnoreCertificateErrors) {
        $params.SkipCertificateCheck = $true
    }

    try {
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        $__logger.Error("Sunshine API Error: $_")
        throw $_
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
    $gamesList.AddRange($installedGames)
    
    $count = Sync-Games -GamesToSync $gamesList -RemoveMissing $true
    $__logger.Info("Sunshine Export: Synced $count games on startup.")
}

function SunshineExport {
    param(
        $scriptMainMenuItemActionArgs
    )

    $shortcutsCreatedCount = Sync-Games -GamesToSync $PlayniteApi.MainView.SelectedGames -RemoveMissing $false
    $PlayniteApi.Dialogs.ShowMessage("Exported $shortcutsCreatedCount games to Sunshine.", "Sunshine Export")
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
        # We need to iterate backwards because removing items changes indices?
        # Wait, Sunshine API uses index. If we remove item at index 0, does item at index 1 become 0?
        # Most likely yes. So we should verify this or iterate backwards.
        # However, the API documentation says: "Delete an application." DELETE /api/apps/{index}
        # It's safer to iterate backwards.
        
        # Also, we need to be careful: if we remove an app, the indices of subsequent apps shift.
        # So we should probably re-fetch the list or be very careful.
        # Re-fetching is safer but slower.
        # Iterating backwards is usually safe for removal by index.
        
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
                            $__logger.Error("Failed to remove Sunshine app at index $i: $_")
                        }
                    }
                }
            }
        }
    }

    return $shortcutsCreatedCount
}