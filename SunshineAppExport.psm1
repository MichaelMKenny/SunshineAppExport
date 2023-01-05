function GetMainMenuItems {
    param(
        $getMainMenuItemsArgs
    )

    $menuItem1 = New-Object Playnite.SDK.Plugins.ScriptMainMenuItem
    $menuItem1.Description = "Export selected games"
    $menuItem1.FunctionName = "SunshineExport"
    $menuItem1.MenuSection = "@Sunshine App Export"
    
    return $menuItem1
}

function SunshineExport {
    param(
        $scriptMainMenuItemActionArgs
    )

    # Load assemblies
    Add-Type -AssemblyName System.Drawing
    $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
    
    # Set paths
    $appAssetsPath = Join-Path -Path $env:LocalAppData -ChildPath "Sunshine Playnite App Export\Apps"
    if (!(Test-Path $appAssetsPath -PathType Container)) {
        New-Item -ItemType Container -Path $appAssetsPath -Force
    }

    # Set creation counter
    $shortcutsCreatedCount = 0

    $appsPath = "$Env:ProgramW6432\Sunshine\apps.json"
    $json = ConvertFrom-Json (Get-Content $appsPath -Raw)

    foreach ($game in $PlayniteApi.MainView.SelectedGames) {
        $gameLaunchURI = 'playnite://playnite/start/' + "$($game.id)"
        $gameName = $($game.name).Split([IO.Path]::GetInvalidFileNameChars()) -join ''

        # Set cover path and create blank file
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.id, "box-art.png")
        if (!(Test-Path $sunshineGameCoverPath -PathType Container)) {
            New-Item -ItemType File -Path $sunshineGameCoverPath -Force
        }

        $logOutput = [System.IO.Path]::Combine($appAssetsPath, $game.id, "app.log")

        if ($null -ne $game.CoverImage) {

            $sourceCover = $PlayniteApi.Database.GetFullFilePath($game.CoverImage)
            if (($game.CoverImage -notmatch "^http") -and (Test-Path $sourceCover -PathType Leaf)) {

                if ([System.IO.Path]::GetExtension($game.CoverImage) -eq ".png") {
                    Copy-Item $sourceCover $sunshineGameCoverPath -Force
                } else {
                    # Convert cover image to compatible PNG image format
                    try {
                        $image = [System.Drawing.Image]::FromFile($PlayniteApi.Database.GetFullFilePath($game.CoverImage))
                        $image.Save($sunshineGameCoverPath, $imageFormat::png)
                        $image.Dispose()
                    } catch {
                        $image.Dispose()
                        $errorMessage = $_.Exception.Message
                        $__logger.Info("Error converting cover image of `"$($game.name)`". Error: $errorMessage")
                    }
                }

                $newApp = New-Object -TypeName psobject
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "name" -Value $game.name
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "output" -Value $logOutput
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "cmd" -Value $gameLaunchURI
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "image-path" -Value $sunshineGameCoverPath

                $json.apps = $json.apps | ForEach-Object {
                    if ($_.cmd -eq $gameLaunchURI) {
                        $newApp
                    } else {
                        $_
                    } 
                }

                if (!($json.apps | Where-Object { $_.cmd -eq $gameLaunchURI })) {
                    $json.apps += $newApp
                }
            }
        }

        $shortcutsCreatedCount++
    }

    ConvertTo-Json $json -Depth 100 | Out-File $appsPath -Encoding utf8

    # Show finish dialogue with shortcut creation count
    $PlayniteApi.Dialogs.ShowMessage(("Sunshine app shortcuts created: {0}" -f $shortcutsCreatedCount), "Sunshine App Export")
}