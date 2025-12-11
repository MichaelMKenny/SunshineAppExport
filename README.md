# SunshineAppExport

This is a [Playnite](https://github.com/JosefNemec/Playnite) addon that automatically syncs your installed games to [Sunshine](https://github.com/LizardByte/Sunshine), allowing you to stream them via [Moonlight](https://github.com/moonlight-stream).

## Features
- **Automatic Sync**: Syncs all installed games to Sunshine when Playnite starts.
- **Real-Time Sync**: Automatically adds or removes games from Sunshine when they are installed or uninstalled in Playnite.
- **Manual Export**: Export specific selected games via the Extensions menu.
- **Smart Removal**: Automatically removes games from Sunshine when they are uninstalled from Playnite (only affects games created by this extension).
- **Cover Art**: Uploads Playnite box art to Sunshine.

## Configuration
1. Go to **Extensions** -> **Sunshine App Export** -> **Configure Sunshine Export**.
2. Enter your Sunshine API URL (default: `https://localhost:47990`).
3. Enter your Sunshine **Username** and **Password**.
4. (Optional) Check **Ignore Certificate Errors** if you are using a self-signed certificate (common for default Sunshine installs).
5. (Optional) Check **Sync on Playnite Startup** to sync your entire library completely every time Playnite starts (default: Off).
6. (Optional) Check **Keep up to date (Real-time sync)** to automatically add/remove games in Sunshine as you install/uninstall them in Playnite (default: Off).

## Usage
- **Manual**: Select games in Playnite, then select **Extensions** -> **Sunshine App Export** -> **Export selected games**.
- **Automatic**: If enabled in settings, just launch Playnite. Your library will be synced in the background.

## Requirements
- Playnite
- Sunshine (running and accessible)
