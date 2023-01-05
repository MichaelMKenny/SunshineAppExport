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

    # Set creation counter
    $shortcutsCreatedCount = 0

    $appsPath = "$Env:ProgramW6432\Sunshine\apps.json"
    $json = ConvertFrom-Json (Get-Content $appsPath -Raw)

    foreach ($game in $PlayniteApi.MainView.SelectedGames) {
        $gameLaunchURI = 'playnite://playnite/start/' + "$($game.id)"

        $logOutput = "$($game.name.ToLower().Replace(' ', '_')).log"

        if ($null -ne $game.CoverImage) {

            $sourceCover = $PlayniteApi.Database.GetFullFilePath($game.CoverImage)
            if (($game.CoverImage -notmatch "^http") -and (Test-Path $sourceCover -PathType Leaf)) {

                $newApp = New-Object -TypeName psobject
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "name" -Value $game.name
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "output" -Value $logOutput
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "cmd" -Value $gameLaunchURI
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "image-path" -Value $sourceCover

                $json.apps += $newApp
            }
        }

        $shortcutsCreatedCount++
    }

    ConvertTo-Json $json -Depth 100 | Out-File $appsPath -Encoding utf8

    # Show finish dialogue with shortcut creation count
    $PlayniteApi.Dialogs.ShowMessage(("Sunshine app shortcuts created: {0}" -f $shortcutsCreatedCount), "Sunshine App Export")
}