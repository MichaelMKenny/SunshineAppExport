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
            $__logger.Info("Sunshine DEBUG: Request Body: $json")
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
    # Only set startup time - actual sync is deferred to OnLibraryUpdated
    # This avoids race conditions with lazy-loading libraries like Epic Games Store
    $script:AppStartTime = Get-Date
    $script:HasRunStartupSync = $false
    $__logger.Info("Sunshine Export: Application started, sync will run after library update.")
}

function OnLibraryUpdated {
    $config = Get-PluginConfig
    if ($config.SyncOnStartup) {
        # Only run once during startup window, use flag to prevent duplicates
        if ($script:HasRunStartupSync) {
            $__logger.Info("Sunshine Export: Startup sync already completed, skipping.")
            return
        }
        
        # Allow syncs triggered by library updates for the first 2 minutes to catch lazy-loaded libraries (like Epic)
        if ($script:AppStartTime -and (Get-Date) -lt $script:AppStartTime.AddMinutes(2)) {
            # Mark as running immediately to prevent concurrent syncs
            $script:HasRunStartupSync = $true
            
            $__logger.Info("Sunshine Export: Triggering startup sync due to library update...")
            
            # Wait 10 seconds to ensure ALL library plugins have fully propagated IsInstalled states
            # Epic Games Store is particularly slow to report installed status
            Start-Sleep -Seconds 10

            $installedGames = $PlayniteApi.Database.Games | Where-Object { $_.IsInstalled }
            $gamesList = New-Object System.Collections.Generic.List[Playnite.SDK.Models.Game]
            if ($installedGames) {
                foreach ($game in $installedGames) {
                    $gamesList.Add($game)
                }
            }
            
            $count = Sync-Games -GamesToSync $gamesList -RemoveMissing $true
            $__logger.Info("Sunshine Export: Synced $count games after library update.")
        }
    }
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

function PrepareGameCover {
    param(
        [Playnite.SDK.Models.Game]$Game,
        [string]$TargetPath
    )
    
    # Ensure target directory exists
    $targetDir = [System.IO.Path]::GetDirectoryName($TargetPath)
    if (!(Test-Path $targetDir -PathType Container)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }
    
    # Create blank file if it doesn't exist
    if (!(Test-Path $TargetPath -PathType Leaf)) {
        New-Item -ItemType File -Path $TargetPath -Force | Out-Null
    }

    if ($null -ne $Game.CoverImage) {
        $sourceCover = $PlayniteApi.Database.GetFullFilePath($Game.CoverImage)
        if (($Game.CoverImage -notmatch "^http") -and (Test-Path $sourceCover -PathType Leaf)) {
            if ([System.IO.Path]::GetExtension($Game.CoverImage) -eq ".png") {
                Copy-Item $sourceCover $TargetPath -Force
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
                
                    $fileStream = New-Object System.IO.FileStream($TargetPath, [System.IO.FileMode]::Create)
                    $encoder.Save($fileStream)
                    $fileStream.Close()
                }
                catch {
                    if ($null -ne $bitmap) { $bitmap = $null }
                    if ($null -ne $fileStream) { $fileStream.Close() }
                    $errorMessage = $_.Exception.Message
                    $__logger.Info("Error converting cover image of `"$($Game.Name)`". Error: $errorMessage")
                }
            }
        }
    }
    
    # Verify cover image exists and is valid (not 0 bytes)
    $finalCoverPath = $TargetPath
    try {
        $coverItem = Get-Item $TargetPath -ErrorAction Stop
        if ($coverItem.Length -eq 0) {
            $finalCoverPath = ""
        }
    }
    catch {
        $finalCoverPath = ""
    }
    
    return $finalCoverPath
}

function Sync-Games {
    param(
        [System.Collections.Generic.List[Playnite.SDK.Models.Game]]$GamesToSync,
        [bool]$RemoveMissing = $false
    )

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

    # --- Phase 1: Deduplicate Existing Apps ---
    # We must do this before syncing because deleting items shifts indices, which invalidates logic if done inline.
    $duplicatesFound = $false
    $idsToIndexMap = @{} # Key: GameID, Value: List of Indices

    # 1. Map all existing apps
    for ($i = 0; $i -lt $sunshineApps.Count; $i++) {
        $app = $sunshineApps[$i]
        if ($app.detached -and $app.detached.Length -gt 0) {
            $existingGameId = GetGameIdFromCmd($app.detached[0])
            if (![string]::IsNullOrEmpty($existingGameId)) {
                if (-not $idsToIndexMap.ContainsKey($existingGameId)) {
                    $idsToIndexMap[$existingGameId] = New-Object System.Collections.Generic.List[int]
                }
                $idsToIndexMap[$existingGameId].Add($i)
            }
        }
    }

    # 2. Convert to flat list of indices to remove (Keep the first/lowest index, remove others)
    $indicesToRemove = New-Object System.Collections.Generic.List[int]
    foreach ($key in $idsToIndexMap.Keys) {
        $indices = $idsToIndexMap[$key]
        if ($indices.Count -gt 1) {
            # Keep index 0 (the first one found, usually lowest), remove the rest
            for ($k = 1; $k -lt $indices.Count; $k++) {
                $indicesToRemove.Add($indices[$k])
            }
        }
    }

    # 3. Remove duplicates (Backwards to preserve indices)
    if ($indicesToRemove.Count -gt 0) {
        $duplicatesFound = $true
        $indicesToRemove.Sort() 
        $indicesToRemove.Reverse() # Remove from end to start
        
        foreach ($idx in $indicesToRemove) {
            try {
                Remove-SunshineApp -Index $idx
                $__logger.Info("Sunshine Export: Removed duplicate app at index $idx")
            }
            catch {
                $__logger.Error("Failed to remove duplicate app at index $idx : $_")
            }
        }
    }

    # 4. Re-fetch apps if we modified anything
    if ($duplicatesFound) {
        Start-Sleep -Milliseconds 500 # Brief pause for API to settle
        try {
            $sunshineAppsResponse = Get-SunshineApps
            $sunshineApps = $sunshineAppsResponse.apps
        }
        catch {
            $__logger.Error("Failed to re-fetch apps after deduplication.")
            return 0
        }
    }
    # ------------------------------------------

    # Create a hashset of Game IDs to sync for O(1) lookup
    $gamesToSyncIds = New-Object System.Collections.Generic.HashSet[string]
    $gameNames = @()
    foreach ($game in $GamesToSync) {
        $gamesToSyncIds.Add($game.Id.ToString()) | Out-Null
        $gameNames += $game.Name
    }
    $__logger.Info("Sunshine Export: Found installed games: $($gameNames -join ', ')")

    # Build a map of existing Game IDs to their Sunshine index for fast lookup
    $existingGameIdToIndex = @{}
    for ($i = 0; $i -lt $sunshineApps.Count; $i++) {
        $app = $sunshineApps[$i]
        if ($app.detached -and $app.detached.Length -gt 0) {
            $existingGameId = GetGameIdFromCmd($app.detached[0])
            if (![string]::IsNullOrEmpty($existingGameId)) {
                # Only keep the first occurrence (duplicates handled in Phase 1)
                if (-not $existingGameIdToIndex.ContainsKey($existingGameId)) {
                    $existingGameIdToIndex[$existingGameId] = $i
                }
            }
        }
    }

    # Separate games into NEW (not in Sunshine) and EXISTING (already in Sunshine)
    $newGames = @()
    $existingGames = @()
    foreach ($game in $GamesToSync) {
        if ($existingGameIdToIndex.ContainsKey($game.Id.ToString())) {
            $existingGames += $game
        }
        else {
            $newGames += $game
        }
    }
    
    $__logger.Info("Sunshine Export: $($newGames.Count) new games, $($existingGames.Count) existing games to sync.")

    # --- Phase 2A: Add NEW games first ---
    # This must be done before updates because adding apps can shift indices
    foreach ($game in $newGames) {
        $gameLaunchCmd = "`"$playniteExecutablePath`" --start $($game.Id)"
        
        # Prepare cover image
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.Id, "box-art.png")
        $finalCoverPath = PrepareGameCover -Game $game -TargetPath $sunshineGameCoverPath
        
        $newApp = @{
            name         = $game.Name
            detached     = @($gameLaunchCmd)
            "image-path" = $finalCoverPath
            index        = -1  # New app
        }

        try {
            Add-SunshineApp -App $newApp
            $shortcutsCreatedCount++
            $__logger.Info("Sunshine Export: Added new game: $($game.Name)")
        }
        catch {
            $__logger.Error("Failed to add new app for game $($game.Name): $_")
        }
    }

    # --- Phase 2B: Re-fetch apps after adding new ones to get fresh indices ---
    if ($newGames.Count -gt 0 -and $existingGames.Count -gt 0) {
        Start-Sleep -Milliseconds 300  # Brief pause for API to settle
        try {
            $sunshineAppsResponse = Get-SunshineApps
            $sunshineApps = $sunshineAppsResponse.apps
            
            # Rebuild the index map with fresh data
            $existingGameIdToIndex = @{}
            for ($i = 0; $i -lt $sunshineApps.Count; $i++) {
                $app = $sunshineApps[$i]
                if ($app.detached -and $app.detached.Length -gt 0) {
                    $existingGameId = GetGameIdFromCmd($app.detached[0])
                    if (![string]::IsNullOrEmpty($existingGameId)) {
                        if (-not $existingGameIdToIndex.ContainsKey($existingGameId)) {
                            $existingGameIdToIndex[$existingGameId] = $i
                        }
                    }
                }
            }
        }
        catch {
            $__logger.Error("Failed to re-fetch apps after adding new games: $_")
            # Continue with stale data - updates may be slightly off but better than failing
        }
    }

    # --- Phase 2C: Update EXISTING games with fresh indices ---
    foreach ($game in $existingGames) {
        $gameLaunchCmd = "`"$playniteExecutablePath`" --start $($game.Id)"
        
        # Prepare cover image
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.Id, "box-art.png")
        $finalCoverPath = PrepareGameCover -Game $game -TargetPath $sunshineGameCoverPath
        
        # Get the current index from our map (should exist since we categorized it as existing)
        $currentIndex = $existingGameIdToIndex[$game.Id.ToString()]
        
        $newApp = @{
            name         = $game.Name
            detached     = @($gameLaunchCmd)
            "image-path" = $finalCoverPath
            index        = $currentIndex
        }

        try {
            Add-SunshineApp -App $newApp
            $shortcutsCreatedCount++
            $__logger.Info("Sunshine DEBUG: Updated game at index $currentIndex : $($game.Name)")
        }
        catch {
            $__logger.Error("Failed to update app for game $($game.Name): $_")
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
                            $__logger.Error("Failed to remove Sunshine app at index ${i}: $_")
                        }
                    }
                }
            }
        }
    }

    return $shortcutsCreatedCount
}
