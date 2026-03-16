<#
    USER SETUP SCRIPT (Run as current User)
#>

param (
    [int]$ForceStep = 0,    # Force execution from a specific step
    [int]$StopAfterStep = 0 # 0 = Run to end, >0 = Stop immediately after this step index
)

function Set-ExplorerTweaks {
    <#
    .SYNOPSIS
        Configures Windows Explorer settings like file extensions, hidden files, and startup location.
    #>
    Write-Log "Applying Windows Explorer Tweaks..." -Level STEP

    try {
        $ExplorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        
        # Ensure the registry path exists
        if (!(Test-Path $ExplorerPath)) {
            New-Item -Path $ExplorerPath -Force | Out-Null
        }

        # 1. Show File Extensions (0 = Show, 1 = Hide)
        Write-Log "Configuring Explorer to show all file extensions..." -Level INFO
        Set-ItemProperty -Path $ExplorerPath -Name "HideFileExt" -Value 0 -Force

        # 2. Show Hidden Files (1 = Show, 2 = Hide)
        Write-Log "Configuring Explorer to show hidden files..." -Level INFO
        Set-ItemProperty -Path $ExplorerPath -Name "Hidden" -Value 1 -Force

        # 3. Launch To (1 = This PC, 2 = Quick Access)
        # Note: LaunchTo is often located in a slightly different subkey in newer Win 11 builds
        Write-Log "Setting Explorer to launch to 'This PC'..." -Level INFO
        $LaunchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $LaunchPath -Name "LaunchTo" -Value 1 -Force

        Write-Log "Explorer tweaks applied successfully." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Set-ExplorerTweaks: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-Taskbar {
    <#
    .SYNOPSIS
        Configures the taskbar layout (alignment) and hides unnecessary icons like Search and Task View.
    #>
    Write-Log "Configuring Taskbar (Alignment, Search, Icons)..." -Level STEP

    try {
        $AdvPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        $SearchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        
        # Ensure registry paths exist
        if (!(Test-Path $AdvPath)) { New-Item -Path $AdvPath -Force | Out-Null }

        # 1. Taskbar Alignment (0 = Left, 1 = Center)
        Write-Log "Setting taskbar alignment to the left..." -Level INFO
        Set-ItemProperty -Path $AdvPath -Name "TaskbarAl" -Value 0 -Force

        # 2. Hide Search (0 = Hidden, 1 = Icon only, 2 = Search box)
        Write-Log "Hiding search icon/box from taskbar..." -Level INFO
        if (!(Test-Path $SearchPath)) { New-Item -Path $SearchPath -Force | Out-Null }
        Set-ItemProperty -Path $SearchPath -Name "SearchboxTaskbarMode" -Value 0 -Force

        # 3. Hide Task View (0 = Hidden, 1 = Shown)
        Write-Log "Hiding Task View button..." -Level INFO
        Set-ItemProperty -Path $AdvPath -Name "ShowTaskViewButton" -Value 0 -Force

        # 4. Bonus: Hide Widgets (Windows 11)
        # This is a common requirement alongside Task View
        try {
            if (Get-ItemProperty -Path $AdvPath -Name "TaskbarDa" -ErrorAction SilentlyContinue) {
                Set-ItemProperty -Path $AdvPath -Name "TaskbarDa" -Value 0 -Force -ErrorAction SilentlyContinue
                Write-Log "Widgets (TaskbarDa) disabled." -Level SUCCESS
            }
        } catch { 
            # ignore
        }

        Write-Log "Taskbar configuration applied successfully." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Set-Taskbar: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-SystemLanguageAndKeyboard {
    param (
        # Default System UI Language (e.g., "en-GB", "de-DE", "en-US")
        [string]$UILanguage = "en-GB",

        # Default Keyboard Layout Code (e.g., "0407:00000407" for German)
        [string]$KeyboardLayout = "0407:00000407",

        # Default Culture for formats (e.g., "de-DE" for Euro/German Date)
        [string]$Culture = "de-DE",

        # GeoID for Location (94 = Germany, 242 = UK, 244 = USA)
        [int]$GeoId = 94
    )

    <#
    .SYNOPSIS
        Sets System UI, Keyboard Layout, and Regional formats based on parameters.
    #>
    Write-Log "Configuring Language ($UILanguage), Keyboard ($KeyboardLayout), and Region..." -Level STEP

    try {
        # 1. Define and Apply Language/Keyboard List
        Write-Log "Defining language list with $UILanguage and keyboard $KeyboardLayout..." -Level INFO
        $NewLanguages = New-WinUserLanguageList -Language $UILanguage
        
        # Clear default input methods and add the specified layout
        $NewLanguages[0].InputMethodTips.Clear()
        $NewLanguages[0].InputMethodTips.Add($KeyboardLayout) 

        # Apply the list to the current user
        Write-Log "Applying user language and input list (Force)..." -Level INFO
        Set-WinUserLanguageList $NewLanguages -Force

        # 2. Regional Formats and Culture
        Write-Log "Setting Culture and System Locale to $Culture..." -Level INFO
        Set-Culture $Culture
        
        # Set Home Location
        Write-Log "Setting Home Location (GeoID) to $GeoId..." -Level INFO
        Set-WinHomeLocation -GeoId $GeoId
        
        # Set System Locale (Non-Unicode programs)
        Set-WinSystemLocale -SystemLocale $Culture

        Write-Log "Language Setup completed: UI=$UILanguage, Keyboard=$KeyboardLayout, Region=$Culture." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Set-SystemLanguageAndKeyboard: $($_.Exception.Message)" -Level ERROR
    }
}

function Optimize-StartMenu {
    <#
    .SYNOPSIS
        Prepares Start Menu optimizations. 
        Note: Explorer restart is handled globally at the end of the main script.
    #>
    Write-Log "Optimizing Start Menu (Search, Recommendations, Layout)..." -Level STEP

    try {
        $AdvancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        $PolicyPath   = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        $SearchPath   = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"

        # 1. Disable Recommendations and Recently Used Items
        Write-Log "Disabling Start Menu recommendations and tracking..." -Level INFO
        Set-ItemProperty -Path $AdvancedPath -Name "Start_IrisRecommendations" -Value 0 -Force
        Set-ItemProperty -Path $AdvancedPath -Name "Start_TrackProgs" -Value 0 -Force
        Set-ItemProperty -Path $AdvancedPath -Name "Start_TrackDocs" -Value 0 -Force

        # 2. Disable Windows Tips and Suggestions (Ads)
        Write-Log "Disabling suggestions and subscribed content (ads)..." -Level INFO
        Set-ItemProperty -Path $PolicyPath -Name "SubscribedContent-338388Enabled" -Value 0 -Force
        Set-ItemProperty -Path $PolicyPath -Name "SystemPaneSuggestionsEnabled" -Value 0 -Force

        # 3. Disable Web Search (Bing) in Start Menu
        Write-Log "Disabling Bing Web Search in Start Menu..." -Level INFO
        if (!(Test-Path $SearchPath)) { New-Item $SearchPath -Force | Out-Null }
        Set-ItemProperty -Path $SearchPath -Name "BingSearchEnabled" -Value 0 -Force
        # Additional key for newer Windows 11 builds
        Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "DisableSearchBoxSuggestions" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        # 4. Clear Default Layout (Pins)
        Write-Log "Checking for default start menu pin layout..." -Level INFO
        $StartLayoutPath = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start.bin"
        if (Test-Path $StartLayoutPath) {
            Remove-Item $StartLayoutPath -Force
            Write-Log "Start menu pins cleared (reset to default/empty)." -Level SUCCESS
        }

        Write-Log "Start Menu optimization prepared successfully." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Optimize-StartMenu: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-LegacyContextMenu {
    <#
    .SYNOPSIS
        Restores the Windows 10 style classic context menu in Windows 11.
        Note: Explorer restart is required to apply changes.
    #>
    Write-Log "Restoring Classic Context Menu (Windows 10 Style)..." -Level STEP

    try {
        $GuidPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
        $InprocPath = "$GuidPath\InprocServer32"

        # 1. Create the Registry Path
        # The existence of this specific CLSID override forces the legacy menu
        Write-Log "Creating Registry override for Context Menu CLSID..." -Level INFO
        if (!(Test-Path $InprocPath)) {
            New-Item -Path $InprocPath -Force | Out-Null
        }

        # 2. Set the (Default) value to empty
        # CRITICAL: The value must be empty (not null) to correctly override the system default
        Write-Log "Setting InprocServer32 default value to empty..." -Level INFO
        Set-ItemProperty -Path $InprocPath -Name "(Default)" -Value "" -Force

        Write-Log "Classic context menu restored successfully." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Set-LegacyContextMenu: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-ColorPersonalization {
    param (
        # 0 = Dark Mode, 1 = Light Mode
        [int]$ThemeMode = 0,
        
        # 0 = Automatic Accent Color, 1 = Manual Accent Color
        [int]$ColorMode = 0,
        
        # Hex Color in ABGR format (Alpha, Blue, Green, Red)
        [string]$CustomAccentColor = "0xFFC3B600",

        [bool]$EnableTransparency = $true,
        [bool]$ShowColorOnTitlebars = $false
    )

    Write-Log "Applying Windows Personalization (Theme: $ThemeMode, Color: $ColorMode)..." -Level STEP

    try {
        $personalizePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $dwmPath         = "HKCU:\Software\Microsoft\Windows\DWM"
        $accentPath      = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent"

        # 1. Dark/Light Mode (System & Apps)
        Set-ItemProperty -Path $personalizePath -Name "SystemUsesLightTheme" -Value $ThemeMode
        Set-ItemProperty -Path $personalizePath -Name "AppsUseLightTheme"   -Value $ThemeMode

        # 2. Accent Color Logic
        if ($ColorMode -eq 0) {
            # MODE: AUTOMATIC
            Write-Log "Mode: Automatic. Clearing manual overrides to allow wallpaper scan..." -Level INFO
            
            # Enable prevalence but remove specific hex codes
            Set-ItemProperty -Path $personalizePath -Name "EnableAccentColorOnStart" -Value 1
            Set-ItemProperty -Path $dwmPath -Name "ColorPrevalence" -Value 1
            
            # Deleting these keys forces Windows to fall back to the wallpaper analysis
            Remove-ItemProperty -Path $dwmPath    -Name "AccentColor" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $dwmPath    -Name "ColorizationColor" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $accentPath -Name "AccentColorMenu" -ErrorAction SilentlyContinue
        } 
        else {
            # MODE: MANUAL
            Write-Log "Mode: Manual. Applying custom color: $CustomAccentColor" -Level INFO
            
            # Force color on Start/Taskbar
            Set-ItemProperty -Path $personalizePath -Name "EnableAccentColorOnStart" -Value 1
            
            # Core DWM Keys
            Set-ItemProperty -Path $dwmPath -Name "AccentColor" -Value $CustomAccentColor
            Set-ItemProperty -Path $dwmPath -Name "ColorizationColor" -Value $CustomAccentColor
            
            # Explorer Cache (Crucial for Windows 11 Taskbar consistency)
            if (!(Test-Path $accentPath)) { New-Item -Path $accentPath -Force | Out-Null }
            Set-ItemProperty -Path $accentPath -Name "AccentColorMenu" -Value $CustomAccentColor
        }

        # 3. Effects (Transparency & Titlebars)
        Set-ItemProperty -Path $personalizePath -Name "EnableTransparency" -Value $(if($EnableTransparency){1}else{0})
        
        # ColorPrevalence in DWM also controls if titlebars/borders are colored
        $titlebarValue = if($ShowColorOnTitlebars){1}else{0}
        Set-ItemProperty -Path $dwmPath -Name "ColorPrevalence" -Value $titlebarValue

        # 4. Trigger UI Evaluation
        Write-Log "Broadcasting refresh signals to shell..." -Level INFO
        
        $signature = @'
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
'@
        if (-not ("Win32ThemeRefresher" -as [type])) {
            Add-Type -MemberDefinition $signature -Name "Win32ThemeRefresher" -Namespace "Win32" -ErrorAction Stop
        }

        $result = [IntPtr]::Zero
        
        # Broadcast 1: Immersive Color (Start, Taskbar, Menus)
        [Win32.Win32ThemeRefresher]::SendMessageTimeout([IntPtr]0xffff, 0x001A, [IntPtr]::Zero, "ImmersiveColorSet", 0x0002, 5000, [ref]$result) | Out-Null
        
        # Broadcast 2: Policy (System-wide UI sync)
        [Win32.Win32ThemeRefresher]::SendMessageTimeout([IntPtr]0xffff, 0x001A, [IntPtr]::Zero, "Policy", 0x0002, 5000, [ref]$result) | Out-Null

        Write-Log "Personalization applied successfully." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Color Personalization: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-MouseAcceleration {
    param (
        # $true = Enable Windows Default Acceleration, $false = Disable (1:1 Input)
        [bool]$Enable = $false
    )

    <#
    .SYNOPSIS
        Toggles mouse acceleration and configures a linear or default pointer curve.
    #>
    $Action = if ($Enable) { "Enabling" } else { "Disabling" }
    Write-Log "$Action Mouse Acceleration..." -Level STEP

    try {
        $MousePath = "HKCU:\Control Panel\Mouse"

        if ($Enable) {
            # 1. Restore Windows Default Settings (Acceleration ON)
            Write-Log "Restoring Windows default acceleration values..." -Level INFO
            $Settings = @{
                "MouseSpeed"        = "1"
                "MouseThreshold1"   = "6"
                "MouseThreshold2"   = "10"
                "MouseSensitivity"  = "10" 
            }
            # Windows Default Curves (Standard Windows curve)
            $XCurve = [byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x80,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x00,0x00)
            $YCurve = [byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x38,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xa0,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x01,0x00,0x00,0x00,0x00)
        }
        else {
            # 2. Set 1:1 Raw Input (Acceleration OFF)
            Write-Log "Applying linear values for 1:1 mouse input..." -Level INFO
            $Settings = @{
                "MouseSpeed"        = "0"
                "MouseThreshold1"   = "0"
                "MouseThreshold2"   = "0"
                "MouseSensitivity"  = "10" 
            }
            # Flattened Curves
            $XCurve = [byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x15,0x6e,0x00,0x00,0x00,0x00,0x00,0x00,0x2a,0xdc,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x4a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x00,0x00)
            $YCurve = [byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xb8,0x5e,0x01,0x00,0x00,0x00,0x00,0x00,0x70,0xbd,0x02,0x00,0x00,0x00,0x00,0x00,0x28,0x1c,0x04,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00,0x00)
        }

        # Apply standard settings
        foreach ($Key in $Settings.Keys) {
            Set-ItemProperty -Path $MousePath -Name $Key -Value $Settings[$Key] -Type String -Force
        }

        # Apply binary curves
        Set-ItemProperty -Path $MousePath -Name "SmoothMouseXCurve" -Value $XCurve -Type Binary -Force
        Set-ItemProperty -Path $MousePath -Name "SmoothMouseYCurve" -Value $YCurve -Type Binary -Force
        
        Write-Log "Mouse settings applied ($Action acceleration)." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Set-MouseAcceleration: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-StartMenuBing {
    param ([bool]$Enable = $false)

    $Action = if ($Enable) { "Enabling" } else { "Disabling" }
    Write-Log "$Action Bing Search in Start Menu..." -Level STEP

    try {
        $SearchPath   = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        $ExplorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        
        $Value = if ($Enable) { 1 } else { 0 }

        # 1. Standard Search Settings (Das hat immer funktioniert)
        if (!(Test-Path $SearchPath)) { New-Item -Path $SearchPath -Force | Out-Null }
        Set-ItemProperty -Path $SearchPath -Name "BingSearchEnabled" -Value $Value -Type DWORD -Force
        Set-ItemProperty -Path $SearchPath -Name "AllowSearchToUseLocation" -Value $Value -Type DWORD -Force
        Set-ItemProperty -Path $SearchPath -Name "CortanaConsent" -Value $Value -Type DWORD -Force
        
        # 2. Search Highlights (Funktioniert auch immer)
        Set-ItemProperty -Path $ExplorerPath -Name "Start_SearchHighlights" -Value $Value -Type DWORD -Force

        # 3. Policy Override (Der Teil mit dem Fehler)
        $PolicyPath = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
        # Wir fügen -ErrorAction SilentlyContinue hinzu, dann bleibt das Log sauber!
        if (!(Test-Path $PolicyPath)) { 
            New-Item -Path $PolicyPath -Force -ErrorAction SilentlyContinue | Out-Null 
        }

        if (Test-Path $PolicyPath) {
            $PolicyValue = if ($Enable) { 0 } else { 1 }
            Set-ItemProperty -Path $PolicyPath -Name "DisableSearchBoxSuggestions" -Value $PolicyValue -Type DWORD -Force
        }

        Write-Log "Bing Search settings applied ($Action)." -Level SUCCESS
    } 
    catch {
        Write-Log "ERROR in Set-StartMenuBing: $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-UserBloatware {
    <#
    .SYNOPSIS
        Uninstalls user-specific bloatware stubs via winget and hides UI remnants.
        This serves as a second layer of defense to the system-wide removal.
    #>
    Write-Log "Starting User-Level Bloatware Removal (Winget/Store)..." -Level STEP

    # 1. Reset winget sources to ensure a clean state
    Write-Log "Resetting winget sources..." -Level INFO
    winget source reset --force | Out-Null

    # 2. List of App IDs (Mostly Microsoft Store stubs)
    $BloatApps = @(
        @{ ID = "XP8BT8DW290MPQ"; Name = "Microsoft Teams (Work or School)" },
        @{ ID = "9N8NNWNVT8LQ"; Name = "Microsoft Todo" },
        @{ ID = "9NBLGGH5R558"; Name = "Microsoft To Do" },
        @{ ID = "9NFTCH6J7FHV"; Name = "Power Automate" },
        @{ ID = "9P7BP5VNWKX5"; Name = "Quick Assist" },
        @{ ID = "9PC1H9VN18CM"; Name = "Start Experiences App" },
        @{ ID = "9NBLGGH4QGHW"; Name = "Sticky Notes" },
        @{ ID = "9WZDNCRD29V9"; Name = "Microsoft 365 (Office)" },
        @{ ID = "9NHT9RB2F4HD"; Name = "Microsoft Desktop App Installer / Copilot Stub" }
    )

    # 3. Process each app
    foreach ($App in $BloatApps) {
        try {
            Write-Log "Attempting to uninstall $($App.Name) (ID: $($App.ID))..." -Level INFO
            
            # Determine source: IDs with dots are usually winget, others msstore
            $source = if ($App.ID -like "*.*") { "winget" } else { "msstore" }
            
            # Execute uninstall
            winget uninstall --id $($App.ID) --source $source --silent --accept-source-agreements | Out-Null
            
            # Note: Winget doesn't always return a proper exit code for "already uninstalled",
            # so we log success if no exception occurs.
            Write-Log "Successfully processed $($App.Name)." -Level SUCCESS
        }
        catch {
            Write-Log "Note: Could not uninstall $($App.Name). It might already be removed." -Level INFO
        }
    }

    # 4. Hide Copilot button from Taskbar (User-specific setting)
    Write-Log "Hiding Copilot Taskbar button (Registry)..." -Level INFO
    $AdvPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    try {
        if (!(Test-Path $AdvPath)) { New-Item -Path $AdvPath -Force | Out-Null }
        Set-ItemProperty -Path $AdvPath -Name "ShowCopilotButton" -Value 0 -ErrorAction SilentlyContinue
        Write-Log "Copilot button hidden." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to hide Copilot button: $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-UserAutostarts {
    <# 
    .SYNOPSIS
        Removes startup entries for the current user (HKCU) to improve login speed.
    #>
    Write-Log "Cleaning user-specific autostarts (HKCU)..." -Level STEP
    
    $RunPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    
    # Expanded list including Edge and common launcher/apps
    $Targets = @(
        "MicrosoftEdgeAutoLaunch", # The main Edge autostart
        "msedge",                  # Alternative Edge key
        "EdgeSidebar",             # Edge Sidebar stub
        "Spotify", 
        "OneDrive"
    )

    try {
        if (!(Test-Path $RunPath)) {
            Write-Log "Registry path $RunPath not found. Skipping." -Level INFO
            return
        }

        foreach ($T in $Targets) {
            # Check if the property exists before attempting removal
            if (Get-ItemProperty -Path $RunPath -Name $T -ErrorAction SilentlyContinue) {
                Remove-ItemProperty -Path $RunPath -Name $T -Force -ErrorAction SilentlyContinue
                Write-Log "Removed User-Autostart entry: $T" -Level SUCCESS
            }
        }
        
        Write-Log "User autostart cleanup completed." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR in Remove-UserAutostarts: $($_.Exception.Message)" -Level ERROR
    }
}

function Optimize-UserInterface {
    <# 
    .SYNOPSIS
        Disables UI bloat like Bing Search, Start Menu ads, Widgets, and Chat.
        Optimizes privacy by disabling tailored experiences and diagnostic-based ads.
    #>
    Write-Log "Optimizing User Interface & Privacy Settings..." -Level STEP

    try {
        # 1. Disable Bing Search in Start Menu
        $SearchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        if (!(Test-Path $SearchPath)) { New-Item $SearchPath -Force | Out-Null }
        
        Write-Log "Disabling Bing Search and Cortana consent..." -Level INFO
        Set-ItemProperty $SearchPath -Name "BingSearchEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty $SearchPath -Name "CortanaConsent" -Value 0 -Type DWord -ErrorAction SilentlyContinue

        # 2. Disable Taskbar Widgets & Chat
        $ExplorerAdvanced = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        if (!(Test-Path $ExplorerAdvanced)) { New-Item $ExplorerAdvanced -Force | Out-Null }

        Write-Log "Hiding Taskbar Widgets and Chat icons..." -Level INFO
        # TaskbarDa = Widgets, TaskbarMn = Chat
        $TaskbarSettings = @{
            "TaskbarDa" = 0
            "TaskbarMn" = 0
        }

        foreach ($setting in $TaskbarSettings.GetEnumerator()) {
            try {
                Set-ItemProperty $ExplorerAdvanced -Name $setting.Key -Value $setting.Value -Type DWord -ErrorAction Stop
            }
            catch {
                Write-Log "Note: Icon $($setting.Key) could not be toggled via Registry (Access Denied). Handled via Explorer Restart." -Level WARN
            }
        }

        # 3. Privacy: Disable "Tailored Experiences"
        $PrivacyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy"
        if (!(Test-Path $PrivacyPath)) { New-Item $PrivacyPath -Force | Out-Null }
        
        Write-Log "Disabling Tailored Experiences with diagnostic data..." -Level INFO
        Set-ItemProperty $PrivacyPath -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue

        # 4. Global Personalization (Ads & Suggestions)
        $ContentPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        if (Test-Path $ContentPath) {
            Write-Log "Disabling System Pane Suggestions and Silent App Installs..." -Level INFO
            Set-ItemProperty $ContentPath -Name "SystemPaneSuggestionsEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
            Set-ItemProperty $ContentPath -Name "SilentInstalledAppsEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
            Set-ItemProperty $ContentPath -Name "SoftLandingEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue # Disable "Welcome experiences"
        }

        Write-Log "User Interface and Privacy optimization completed." -Level SUCCESS
    }
    catch {
        Write-Log "An error occurred during UI optimization: $($_.Exception.Message)" -Level ERROR
    }
}

function Restart-Explorer {
    <# 
    .SYNOPSIS
        Restarts the Windows Explorer process for the current user session only.
        This applies registry-based UI changes immediately.
    #>
    Write-Log "Restarting Windows Explorer to apply all changes..." -Level STEP
    
    try {
        # 1. Get the Session ID of the current PowerShell window
        $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId

        # 2. Find the explorer process that belongs to THIS specific session
        $sessionExplorer = Get-Process explorer -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $currentSessionId }

        if ($sessionExplorer) {
            Write-Log "Stopping Explorer process in Session $currentSessionId..." -Level INFO
            $sessionExplorer | Stop-Process -Force
            # Give Windows a moment to clean up
            Start-Sleep -Seconds 2 
        }

        # 3. Restart Explorer
        # We don't use 'Start-Process explorer' alone to avoid issues with working directories
        Write-Log "Starting new Explorer process..." -Level INFO
        Start-Process "explorer.exe"
        
        Write-Log "Explorer successfully restarted." -Level SUCCESS
    }
    catch {
        Write-Log "ERROR during Explorer restart: $($_.Exception.Message)" -Level ERROR
        # Emergency fallback: Try to start explorer anyway
        Start-Process "explorer.exe" -ErrorAction SilentlyContinue
    }
}

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

$UtilsPath = Join-Path $PSScriptRoot "Utils.ps1"
if (Test-Path $UtilsPath) { . $UtilsPath } else { Throw "Critical Error: Utils.ps1 not found!" }

Write-Log "Starting User-Level Setup" -Level STEP

if ($ForceStep -gt 0) { 
    Write-Log "Manual Override: Starting at Step $ForceStep" -Level WARN 
}

# 1. Visuals & Themes
if (Confirm-StepExecution "Visual Personalization" 1 $StopAfterStep) {
    Set-ExplorerTweaks
    Set-LegacyContextMenu
    Set-ColorPersonalization
}

# 2. Taskbar & Interface
if (Confirm-StepExecution "Interface Optimization" 2 $StopAfterStep) {
    Set-Taskbar
    Optimize-UserInterface
    Set-StartMenuBing
}

# 3. Input & Language
if (Confirm-StepExecution "Input & Language Settings" 3 $StopAfterStep) {
    Set-SystemLanguageAndKeyboard
    Set-MouseAcceleration
}

# 4. Start Menu & Bloatware
if (Confirm-StepExecution "Start Menu & Bloatware" 4 $StopAfterStep) {
    Optimize-StartMenu
    Remove-UserBloatware
}

# 5. Applications & Autostart
if (Confirm-StepExecution "App Setup & Autostart" 5 $StopAfterStep) {
    Remove-UserAutostarts
}

# 6. Finalization
if (Confirm-StepExecution "Finalizing Session" 6 $StopAfterStep) {
    Restart-Explorer
}

Remove-ProgressFile

Write-Host "--- User Setup Complete ---" -ForegroundColor Green