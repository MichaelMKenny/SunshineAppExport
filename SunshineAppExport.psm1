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
    
    return $menuItem1
}

function SunshineExport {
    param(
        $scriptMainMenuItemActionArgs
    )

    $windowCreationOptions = New-Object Playnite.SDK.WindowCreationOptions
    $windowCreationOptions.ShowMinimizeButton = $false
    $windowCreationOptions.ShowMaximizeButton = $false
    
    $window = $PlayniteApi.Dialogs.CreateWindow($windowCreationOptions)
    $window.Title = "Sunshine App Export"
    $window.SizeToContent = "WidthAndHeight"
    $window.ResizeMode = "NoResize"
    
    # Set content of a window. Can be loaded from xaml, loaded from UserControl or created from code behind
    [xml]$xaml = @"
<UserControl
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

    <UserControl.Resources>
        <Style TargetType="TextBlock" BasedOn="{StaticResource BaseTextBlockStyle}" />
    </UserControl.Resources>

    <StackPanel Margin="16,16,16,16">
        <TextBlock Text="Enter your Sunshine apps.json path" 
        Margin="0,0,0,8" 
        VerticalAlignment="Center"/>

        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBox x:Name="SunshinePath"
            Grid.Column="0"
            Margin="0,0,8,0"
            Width="320"
            Height="36"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Top"/>

            <Button x:Name="BrowseButton"
            Grid.Column="1"
            Width="72"
            Height="36"
            HorizontalAlignment="Right"
            Content="Browse"/>
        </Grid>

        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <CheckBox x:Name="SunshineAttached"
            Grid.Column="0"
            Margin="0,0,8,0"
            Width="36"
            Height="36"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Top"
            IsChecked="True"/>

            <TextBlock Text="Close game when SunShine session ends"
            Grid.Column="1"
            VerticalAlignment="Center"
            HorizontalAlignment="Left"/>
        </Grid>

        <TextBlock x:Name="FinishedMessage" 
        Margin="0,0,0,24" 
        VerticalAlignment="Center"
        Visibility="Collapsed"/>

        <Button x:Name="OKButton"
        IsDefault="true"
        Height= "36"
        Margin="0,5,0,0"
        HorizontalAlignment="Center"
        Content="Export Games"/>
    </StackPanel>
</UserControl>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window.Content = [Windows.Markup.XamlReader]::Load($reader)

    $appsPath = "$Env:ProgramW6432\Sunshine\config\apps.json"

    $inputField = $window.Content.FindName("SunshinePath")
    $inputField.Text = $appsPath
    
    # Attach a click event handler to the Browse button
    $browseButton = $window.Content.FindName("BrowseButton")
    $browseButton.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = Split-Path $inputField.Text -Parent
            $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $openFileDialog.Title = "Open apps.json"
    
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $inputField.Text = $openFileDialog.FileName
            }
        })
    

    # Attach a click event handler
    $button = $window.Content.FindName("OKButton")
    $button.Add_Click({
            if ($button.Content -eq "Dismiss") {
                $window.Close()
            }
            else {
                $appsPath = $inputField.Text
                $appsPath = $appsPath -replace '"', ''

                $gameAttached = $window.Content.FindName("SunshineAttached").IsChecked

                $shortcutsCreatedCount = DoWork $appsPath $gameAttached
                $button.Content = "Dismiss"

                $finishedMessage = $window.Content.FindName("FinishedMessage")
                $finishedMessage.Text = ("Created {0} Sunshine app shortcuts" -f $shortcutsCreatedCount)
                $finishedMessage.Visibility = "Visible"
            }
        })
    
    # Set owner if you need to create modal dialog window
    $window.Owner = $PlayniteApi.Dialogs.GetCurrentAppWindow()
    $window.WindowStartupLocation = "CenterOwner"
    
    # Use Show or ShowDialog to show the window
    $window.ShowDialog()
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

function DoWork([string]$appsPath,[bool]$gameAttached) {
    # Load assemblies
    Add-Type -AssemblyName System.Drawing
    $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
    
    # Set paths
    $playniteExecutablePath = Join-Path -Path $PlayniteApi.Paths.ApplicationPath -ChildPath "Playnite.DesktopApp.exe"
    $appAssetsPath = Join-Path -Path $env:LocalAppData -ChildPath "Sunshine Playnite App Export\Apps"
    if (!(Test-Path $appAssetsPath -PathType Container)) {
        New-Item -ItemType Container -Path $appAssetsPath -Force
    }

    # Set creation counter
    $shortcutsCreatedCount = 0

    $json = ConvertFrom-Json (Get-Content $appsPath -Raw)

    foreach ($game in $PlayniteApi.MainView.SelectedGames) {
        $gameLaunchCmd = "`"$playniteExecutablePath`" --start $($game.id)"

        # Set cover path and create blank file
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.id, "box-art.png")
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
                        # Create a BitmapImage and load the JPG image
                        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                        $bitmap.BeginInit()
                        $bitmap.UriSource = New-Object System.Uri($sourceCover, [System.UriKind]::Absolute)
                        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bitmap.EndInit()
                    
                        # Create a PngBitmapEncoder and add the BitmapFrame
                        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
                        $frame = [System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap)
                        $encoder.Frames.Add($frame)
                    
                        # Save the PNG image
                        $fileStream = New-Object System.IO.FileStream($sunshineGameCoverPath, [System.IO.FileMode]::Create)
                        $encoder.Save($fileStream)
                        $fileStream.Close()
                    }
                    catch {
                        if ($null -ne $bitmap) { $bitmap = $null }
                        if ($null -ne $fileStream) { $fileStream.Close() }
                        $errorMessage = $_.Exception.Message
                        $__logger.Info("Error converting cover image of `"$($game.name)`". Error: $errorMessage")
                    }
                }

                $ids = @()
                foreach ($app in $json.apps) {
                    if ($app.id) {
                        $ids += $app.id
                    }
                }

                $id = Get-Random
                while ($ids.Contains($id.ToString())) {
                    $id = Get-Random
                }

                $newApp = New-Object -TypeName psobject
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "name" -Value $game.name
                if( !$gameAttached ) {
                    Add-Member -InputObject $newApp -MemberType NoteProperty -Name "detached" -Value @($gameLaunchCmd)
                }
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "image-path" -Value $sunshineGameCoverPath
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "id" -Value $id.ToString()
                if( $gameAttached ) {
                    Add-Member -InputObject $newApp -MemberType NoteProperty -Name "cmd" -Value $gameLaunchCmd
                }

                $json.apps = $json.apps | ForEach-Object {
                    $found = $false
                    if ($_.detached) {
                        $gameId = GetGameIdFromCmd($_.detached[0])
                        if ($gameId -eq $game.id) {
                            $found = $true
                        }
                    }
                    if( $_.cmd ) {
                        $gameId = GetGameIdFromCmd($_.cmd)
                        if( $gameId -eq $game.id ) {
                            $found = $true
                        }
                    }
                    if( $found ) {
                        $newApp
                    } else {
                        $_
                    }
                }

                if (!($json.apps | Where-Object { 
                    $found = $false
                            if ($_.detached) {
                                $gameId = GetGameIdFromCmd($_.detached[0])
                                if( $gameId -eq $game.id ) {
                                    $found = $true
                                }
                            }
                            if( $_.cmd ) {
                                $gameId = GetGameIdFromCmd($_.cmd)
                                if( $gameId -eq $game.id ) {
                                    $found = $true
                                }
                            }
                            return $found
                        })) {
                    [object[]]$json.apps += $newApp
                }
            }
        }

        $shortcutsCreatedCount++
    }

    $jsonObj = ConvertTo-Json $json -Depth 100
    # Write this using utf8-noBOM, which depending on PS version, is not supported.
    # so as a workaround, we'll use WriteAllLines which defaults to utf8-noBOM
    [System.IO.File]::WriteAllLines("$env:TEMP\apps.json", $jsonObj)


    $result = [System.Windows.Forms.MessageBox]::Show("You will be prompted for administrator rights, as Sunshine now requires administrator rights in order to modify the apps.json file.", "Administrator Required", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return 0
    }
    else {
        Start-Process powershell.exe  -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"Copy-Item -Path $env:TEMP\apps.json -Destination '$appsPath'`"" -WindowStyle Hidden
    }


    return $shortcutsCreatedCount
}
