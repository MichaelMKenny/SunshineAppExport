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
                $shortcutsCreatedCount = doWork($appsPath)
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

function doWork([string]$appsPath) {
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
        $gameLaunchCmd = $playniteExecutablePath + " --start " + "$($game.id)"
        $gameName = $($game.name).Split([IO.Path]::GetInvalidFileNameChars()) -join ''

        # Set cover path and create blank file
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.id, "box-art.png")
        if (!(Test-Path $sunshineGameCoverPath -PathType Container)) {
            $discard = New-Item -ItemType File -Path $sunshineGameCoverPath -Force
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
                        $image = [System.Drawing.Image]::FromFile($sourceCover)
                        $image.Save($sunshineGameCoverPath, $imageFormat::png)
                        $image.Dispose()
                    }
                    catch {
                        $image.Dispose()
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
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "detached" -Value @($gameLaunchCmd)
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "image-path" -Value $sunshineGameCoverPath
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "id" -Value $id.ToString()

                $json.apps = $json.apps | ForEach-Object {
                    if ($_.detached) {
                        if ($_.detached[0] -eq $gameLaunchCmd) {
                            $newApp
                        }
                        else {
                            $_
                        }
                    }
                    else {
                        $_
                    }
                }

                if (!($json.apps | Where-Object { 
                            if ($_.detached) {
                                return $_.detached[0] -eq $gameLaunchCmd
                            }
                            else {
                                return $false
                            }
                        })) {
                    $json.apps += $newApp
                }
            }
        }

        $shortcutsCreatedCount++
    }

    ConvertTo-Json $json -Depth 100 | Out-File $env:TEMP\apps.json -Encoding utf8


    $result = [System.Windows.Forms.MessageBox]::Show("You will be prompted for administrator rights, as Sunshine now requires administrator rights in order to modify the apps.json file.", "Administrator Required", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return 0
    }
    else {
        Start-Process powershell.exe  -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"Copy-Item -Path $env:TEMP\apps.json -Destination '$appsPath'`"" -WindowStyle Hidden
    }
    


    return $shortcutsCreatedCount
}