<#
    SYSTEM SETUP SCRIPT (Run as Administrator)
#>

param (
    [int]$ForceStep = 0,    # Force execution from a specific step
    [int]$StopAfterStep = 0 # 0 = Run to end, >0 = Stop immediately after this step index
)

function Clear-WindowsUpdateCache {
    <#
    .SYNOPSIS
        Purges the Windows Update cache using Robocopy to bypass MAX_PATH (260 character) limitations.
    #>
    Write-Log "Purging Windows Update Cache..." -Level STEP

    $Services = @("wuauserv", "bits", "dosvc")
    $cachePath = "C:\Windows\SoftwareDistribution\Download"
    $emptyDir = Join-Path $env:TEMP "EmptyDirForPurge"

    try {
        # 1. Stop Services
        foreach ($Service in $Services) {
            if ((Get-Service -Name $Service -ErrorAction SilentlyContinue).Status -eq 'Running') {
                Write-Log "Stopping $Service..." -Level INFO
                Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue
            }
        }

        # 2. Use Robocopy to purge the folder (The "Nuclear Option" for long paths)
        if (Test-Path $cachePath) {
            Write-Log "Executing Robocopy purge on $cachePath..." -Level INFO
            
            # Create a temporary empty directory
            if (!(Test-Path $emptyDir)) { New-Item $emptyDir -ItemType Directory -Force | Out-Null }
            
            # /PURGE deletes everything in the destination that isn't in the source (empty)
            # /NJH /NJS /NDL /NC /NS hides the spammy robocopy logs
            robocopy $emptyDir $cachePath /PURGE /NJH /NJS /NDL /NC /NS /MT:16 | Out-Null
            
            # Cleanup the empty helper dir
            Remove-Item $emptyDir -Force -Recurse -ErrorAction SilentlyContinue
            
            Write-Log "Update cache cleared successfully." -Level SUCCESS
        }
    } catch {
        Write-Log "Failed to clear Update Cache: $($_.Exception.Message)" -Level ERROR
    }finally {
        # 3. Restart Services
        Write-Log "Restarting update services..." -Level INFO
        foreach ($Service in $Services) {
            Set-Service -Name $Service -StartupType Manual -ErrorAction SilentlyContinue
            Start-Service -Name $Service -ErrorAction SilentlyContinue
        }

        if (Test-Path $emptyDir) { Remove-Item $emptyDir -Force -Recurse -ErrorAction SilentlyContinue }
        Write-Log "Update services are back online." -Level SUCCESS
    }
}

function Disable-AutomaticDriverInstallation {
    <#
    .SYNOPSIS
        Prevents Windows Update from automatically downloading and installing hardware drivers.
        Includes Registry, Policy, and Metadata settings.
    #>
    
    # Check for Admin rights
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log "Please run this script as Administrator!" -Level ERROR
        return
    }

    Write-Log "Configuring Driver Update Policy..." -Level STEP

    try {
        # 1. Disable Driver Searching (corresponds to sysdm.cpl "No" setting)
        $dsPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching"
        if (!(Test-Path $dsPath)) { New-Item -Path $dsPath -Force | Out-Null }
        Set-ItemProperty -Path $dsPath -Name "SearchOrderConfig" -Value 0 -Type DWord -ErrorAction Stop
        Write-Log "Set SearchOrderConfig to 0 (Manual / sysdm.cpl equivalent)." -Level INFO

        # 2. Prevent drivers in Quality Updates (Registry & Group Policy equivalent)
        $policyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        if (!(Test-Path $policyPath)) { New-Item -Path $policyPath -Force | Out-Null }
        Set-ItemProperty -Path $policyPath -Name "ExcludeWUDriversInQualityUpdate" -Value 1 -Type DWord -ErrorAction Stop
        Write-Log "Set ExcludeWUDriversInQualityUpdate to 1 (Registry/GPEDit equivalent)." -Level INFO

        # 3. Disable Device Metadata (Prevents icons/manufacturer info downloads)
        $metadataPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata"
        if (!(Test-Path $metadataPath)) { New-Item -Path $metadataPath -Force | Out-Null }
        Set-ItemProperty -Path $metadataPath -Name "PreventDeviceMetadataFromNetwork" -Value 1 -Type DWord -ErrorAction Stop
        Write-Log "Set PreventDeviceMetadataFromNetwork to 1 (Metadata disabled)." -Level INFO

        Write-Log "Success: Automatic driver installations are now disabled." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to update driver registry keys: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-DeviceManagerFullTransparency {
    <#
    .SYNOPSIS
        Unlocks the full transparency of the Device Manager by setting three 
        critical environment variables at the machine level.
    #>
    Write-Log "Unlocking full Device Manager transparency..." -Level STEP

    try {
        # Dictionary of the three essential variables
        $DevMgrVars = @{
            # 1. Shows ghost devices (disconnected hardware)
            "DEVMGR_SHOW_NONPRESENT_DEVICES" = "1"
            
            # 2. Shows legacy hidden devices (system-internal drivers)
            "DEVMGR_SHOW_HIDDEN_DEVICES"     = "1"
            
            # 3. Unlocks the 'Details' tab for deeper driver analysis
            "DEVMGR_SHOW_DETAILS"            = "1"
        }

        foreach ($Var in $DevMgrVars.Keys) {
            Write-Log "  > Setting $Var to $($DevMgrVars[$Var])..." -Level INFO
            # Set variables permanently at the Machine level
            [Environment]::SetEnvironmentVariable($Var, $DevMgrVars[$Var], [EnvironmentVariableTarget]::Machine)
        }

        Write-Log "Success: All 3 Device Manager variables are now active." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to set Device Manager variables: $($_.Exception.Message)" -Level ERROR
    }
}

function Disable-TelemetryCompletely {
    <#
    .SYNOPSIS
        Purges system telemetry, data collection services, scheduled tasks, and Bing search integration.
    #>
    Write-Log "Purging System Telemetry & Data Collection..." -Level STEP

    try {
        # 1. Stop and disable telemetry-related services
        $Services = @("DiagTrack", "dmwappushservice", "WerSvc")
        foreach ($Service in $Services) {
            if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue
                Set-Service -Name $Service -StartupType Disabled -ErrorAction Stop
                Write-Log "Service '$Service' disabled." -Level INFO
            }
        }

        # 2. Registry hardening (System & User levels)
        $RegistryPaths = @(
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection",
            "HKCU:\Software\Microsoft\Personalization\Settings",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        )
        foreach ($Path in $RegistryPaths) {
            if (!(Test-Path $Path)) { 
                New-Item -Path $Path -Force | Out-Null 
            }
        }

        # Set Telemetry level to 0 (Security only / Disabled)
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -Type DWord -ErrorAction Stop
        
        # Disable Advertising ID
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Value 0 -Type DWord -ErrorAction Stop
        
        # Disable "Tailored Experiences" based on diagnostic data
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0 -Type DWord -ErrorAction Stop
        Write-Log "Telemetry registry keys neutralized." -Level INFO

        # 3. Disable specific Scheduled Tasks responsible for data collection
        $Tasks = @(
            "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser",
            "\Microsoft\Windows\Application Experience\ProgramDataUpdater",
            "\Microsoft\Windows\Autochk\Proxy",
            "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator"
        )
        foreach ($Task in $Tasks) {
            $taskObj = Get-ScheduledTask -TaskName $Task -ErrorAction SilentlyContinue
            if ($taskObj) {
                Disable-ScheduledTask -TaskName $Task -ErrorAction Stop | Out-Null
                Write-Log "Task '$Task' disabled." -Level INFO
            }
        }

        # 4. Disable Bing Search in Start Menu (prevents keystrokes being sent to MS)
        $SearchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        Set-ItemProperty -Path $SearchPath -Name "BingSearchEnabled" -Value 0 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $SearchPath -Name "CortanaConsent" -Value 0 -Type DWord -ErrorAction Stop
        Write-Log "Bing Search integration disabled." -Level INFO

        Write-Log "Telemetry has been neutralized successfully." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to complete telemetry purge: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-LegacyPasswordLogin {
    <#
    .SYNOPSIS
        Configures Windows to prioritize standard Password login over Windows Hello PIN.
    #>
    Write-Log "Configuring Windows to prioritize Password over PIN..." -Level STEP

    $registryPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Settings\AllowSignInOptions"
    $registryPathHello = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\PasswordLess\Device"

    try {
        # 1. Disable Windows Hello PIN requirement
        # (Equivalent to: Settings > Accounts > Sign-in options > "For improved security, only allow Windows Hello sign-in")
        if (!(Test-Path $registryPathHello)) {
            New-Item -Path $registryPathHello -Force | Out-Null
        }
        Set-ItemProperty -Path $registryPathHello -Name "DevicePasswordLessBuildVersion" -Value 0 -Type DWord -ErrorAction Stop
        Write-Log "Windows Hello PIN requirement disabled." -Level INFO

        # 2. Explicitly allow password sign-in options
        if (!(Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        Set-ItemProperty -Path $registryPath -Name "value" -Value 1 -Type DWord -ErrorAction Stop
        
        Write-Log "Password login is now prioritized." -Level SUCCESS
        Write-Log "Note: You might need to remove an existing PIN manually in 'Sign-in options'." -Level WARN
    }
    catch {
        Write-Log "Failed to set registry keys for password login: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-UACLevel {
    <#
    .SYNOPSIS
        Sets the User Account Control (UAC) to the highest security level (Always Notify).
        Prevents silent elevation of administrative tasks.
    #>
    param (
        [int]$Level = 4 # 4 is 'Always Notify'
    )

    Write-Log "Hardening User Account Control (UAC) to Level: $Level" -Level STEP

    $RegistryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System"

    try {
        # ConsentPromptBehaviorAdmin: 2 = Always notify on secure desktop
        # PromptOnSecureDesktop: 1 = Use the dimmed secure desktop
        $BehaviorValue = if ($Level -eq 4) { 2 } else { 5 } # 5 is default, 2 is max

        Set-ItemProperty -Path $RegistryPath -Name "ConsentPromptBehaviorAdmin" -Value $BehaviorValue -Force -ErrorAction Stop
        Set-ItemProperty -Path $RegistryPath -Name "PromptOnSecureDesktop" -Value 1 -Force -ErrorAction Stop
        
        Write-Log "Success: UAC is now set to 'Always Notify'." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to update UAC registry keys: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-PrivacyHardening {
    <#
    .SYNOPSIS
        Disables Telemetry, Advertising ID, and Data Collection to enhance privacy.
        Sets the Telemetry level to 'Security Only'.
    #>
    Write-Log "Applying Privacy and Telemetry Hardening..." -Level STEP

    $PoliciesPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
    $AdPath       = "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"

    try {
        # 1. Telemetry (0 = Security Only, 1 = Basic, 2 = Enhanced, 3 = Full)
        if (!(Test-Path $PoliciesPath)) { New-Item -Path $PoliciesPath -Force | Out-Null }
        Set-ItemProperty -Path $PoliciesPath -Name "AllowTelemetry" -Value 0 -Force -ErrorAction Stop
        Write-Log "Telemetry set to 'Security Only'." -Level INFO

        # 2. Advertising ID
        if (!(Test-Path $AdPath)) { New-Item -Path $AdPath -Force | Out-Null }
        Set-ItemProperty -Path $AdPath -Name "Enabled" -Value 0 -Force -ErrorAction Stop
        Write-Log "Advertising ID disabled." -Level INFO

        # 3. Disabling 'Tailored Experiences' (Feedback & Diagnostic Data)
        $PrivacyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy"
        if (!(Test-Path $PrivacyPath)) { New-Item -Path $PrivacyPath -Force | Out-Null }
        Set-ItemProperty -Path $PrivacyPath -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0 -Force -ErrorAction SilentlyContinue

        Write-Log "Success: Privacy hardening applied." -Level SUCCESS
    }
    catch {
        Write-Log "Privacy Hardening encountered an error: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-TimeSync {
    <#
    .SYNOPSIS
        Configures the Windows Time Service and synchronizes with a selected NTP source.
    .PARAMETER SourceIndex
        1: PTB (Physikalisch-Technische Bundesanstalt - Highly Recommended for DE)
        2: Google (time.google.com)
        3: Cloudflare (time.cloudflare.com)
        4: Pool NTP (de.pool.ntp.org)
        5: Windows (time.windows.com - Default)
    #>
    param (
        [int]$SourceIndex = 1
    )

    $Sources = @{
        1 = "ptbtime1.ptb.de,0x1"
        2 = "time.google.com,0x1"
        3 = "time.cloudflare.com,0x1"
        4 = "de.pool.ntp.org,0x1"
        5 = "time.windows.com,0x1"
    }

    $SelectedSource = if ($Sources.ContainsKey($SourceIndex)) { $Sources[$SourceIndex] } else { $Sources[5] }
    $ServerName = $SelectedSource.Split(',')[0]

    Write-Log "Configuring Time Sync using Source: $ServerName" -Level STEP

    try {
        # Ensure service is set to start automatically
        Set-Service -Name "W32Time" -StartupType Automatic -ErrorAction Stop

        # Configure NTP Server
        & w32tm /config /manualpeerlist:"$SelectedSource" /syncfromflags:manual /reliable:YES /update 2>$null | Out-Null
        
        # Restart service to apply changes cleanly
        Restart-Service -Name "W32Time" -Force -ErrorAction Stop

        # Trigger immediate re-sync
        Write-Log "  > Triggering immediate resync..." -Level INFO
        & w32tm /resync /nowait 2>$null | Out-Null
        
        Write-Log "Success: System time is now synchronized with $ServerName." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to sync time: $($_.Exception.Message)" -Level WARN
        Write-Log "Check if UDP Port 123 is blocked by a firewall." -Level INFO
    }
}

function Set-FirewallHardening {
    <#
    .SYNOPSIS
        Hardens the Windows Firewall by blocking telemetry, 
        disabling insecure remote rules, and enabling stealth mode.
    #>
    Write-Log "Applying Firewall Hardening..." -Level STEP

    try {
        # 1. Enable Stealth Mode for all profiles
        Write-Log "  > Hardening Firewall profiles..." -Level INFO
        
        # Public & Domain remain strictly hardened
        Set-NetFirewallProfile -Profile Domain, Public `
                            -DefaultInboundAction Block `
                            -DefaultOutboundAction Allow `
                            -AllowUnicastResponseToMulticast False `
                            -ErrorAction SilentlyContinue

        # Private profile allows responses (Required for Home-PC discovery)
        Set-NetFirewallProfile -Profile Private `
                            -DefaultInboundAction Block `
                            -DefaultOutboundAction Allow `
                            -AllowUnicastResponseToMulticast True `
                            -ErrorAction SilentlyContinue

        # 2. Restrict Inbound Rules
        Write-Log "  > Disabling insecure remote rules..." -Level INFO
        $InsecureRules = @("*Remote Assistance*", "*Cast to Device*", "*Distributed Transaction Coordinator*")
        foreach ($RuleName in $InsecureRules) {
            # Check if rule exists before trying to disable to prevent cluttering logs
            if (Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue) {
                Disable-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
            }
        }

        # 3. Outbound Telemetry Block
        Write-Log "  > Creating Outbound Telemetry blocks..." -Level INFO
        $TelemetryApps = @(
            "Telemetry.Api.Internal", 
            "C:\Windows\System32\CompatTelRunner.exe",
            "C:\Windows\System32\DeviceCensus.exe"
        )
        
        foreach ($App in $TelemetryApps) {
            $BaseName = Split-Path $App -Leaf
            $RuleName = "Block_Telemetry_$BaseName"
            
            # Use Name or DisplayName for checking existence
            if (!(Get-NetFirewallRule -DisplayName "Ghost_Block: $RuleName" -ErrorAction SilentlyContinue)) {
                New-NetFirewallRule -DisplayName "Ghost_Block: $RuleName" `
                                    -Direction Outbound `
                                    -Program $App `
                                    -Action Block `
                                    -Enabled True | Out-Null
            }
        }

        Write-Log "Success: Firewall hardening applied." -Level SUCCESS
    }
    catch {
        Write-Log "Firewall Hardening encountered an error: $($_.Exception.Message)" -Level ERROR
    }
}

function Enable-NetworkDiscovery {
    <#
    .SYNOPSIS
        Enables network discovery and file/printer sharing in the firewall
        and starts the required background services.
    #>
    Write-Log "Enabling Network Discovery & File Sharing..." -Level STEP

    try {
        # 1. Enable Firewall rules for Network Discovery and Sharing
        Write-Log "Configuring Firewall rules for discovery groups..." -Level INFO
        Get-NetFirewallRule -DisplayGroup "Network Discovery" | Enable-NetFirewallRule
        Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" | Enable-NetFirewallRule

        # 2. Set required services to 'Automatic' and start them
        # FDResPub: Function Discovery Resource Publication (crucial for visibility)
        # SSDPSRV: SSDP Discovery (finds UPnP devices)
        # UpnPHost: UPnP Device Host
        $Services = @("FDResPub", "SSDPSRV", "upnphost")

        foreach ($Service in $Services) {
            Write-Log "Configuring and starting service: $Service..." -Level INFO
            Set-Service -Name $Service -StartupType Automatic
            Start-Service -Name $Service -ErrorAction SilentlyContinue
        }

        # 3. Optional: SMB 1.0 Support
        # Only enable if you have very legacy hardware (e.g., old NAS). Usually disabled for security.
        # Enable-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol" -NoRestart

        Write-Log "Network Discovery is now ENABLED." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to enable some Network Discovery features: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-NetworkPrivate {
    <#
    .SYNOPSIS
        Sets all active network profiles to 'Private'.
        Ensures local network discovery and printer sharing work within the hardened firewall.
    #>
    Write-Log "Configuring Network Profiles to 'Private'..." -Level STEP

    try {
        $NetworkProfiles = Get-NetConnectionProfile -ErrorAction SilentlyContinue

        foreach ($Profile in $NetworkProfiles) {
            if ($Profile.NetworkCategory -ne "Private") {
                Write-Log "  > Changing Category for '$($Profile.Name)' from $($Profile.NetworkCategory) to Private..." -Level INFO
                Set-NetConnectionProfile -InterfaceIndex $Profile.InterfaceIndex -NetworkCategory Private -ErrorAction Stop
            } else {
                Write-Log "  > Interface '$($Profile.Name)' is already set to Private." -Level INFO
            }
        }
        
        Write-Log "Success: All active networks are now trusted (Private)." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to update network profiles: $($_.Exception.Message)" -Level WARN
    }
}

function Set-HighPerformanceMode {
    <#
    .SYNOPSIS
        Activates the High Performance power plan, disables hibernation, 
        and turns off Fast Boot for maximum system stability and clean driver initialization.
    #>
    Write-Log "Optimizing Power and Performance Settings..." -Level STEP

    # Standard GUID for High Performance
    $highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
    
    try {
        # 1. Activate High Performance Power Plan
        Write-Log "  > Activating High Performance Plan via powercfg..." -Level INFO
        & powercfg /setactive $highPerfGuid
        
        # 2. Disable Hibernation
        # This deletes hiberfil.sys, saving significant SSD space
        Write-Log "  > Disabling Hibernation..." -Level INFO
        & powercfg /hibernate off

        # 3. Disable Windows Fast Boot
        # Prevents the kernel from entering a "hybrid sleep" state, ensuring a clean driver start on every boot
        Write-Log "  > Disabling Windows Fast Boot..." -Level INFO
        $PowerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        if (Test-Path $PowerPath) {
            Set-ItemProperty -Path $PowerPath -Name "HiberbootEnabled" -Value 0 -Force -ErrorAction SilentlyContinue
        }

        Write-Log "Success: Performance mode activated and storage optimized." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to apply power settings: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-CPULatencyTweaks {
    <#
    .SYNOPSIS
        Reduces system-wide latency by adjusting the I/O priority and responsiveness.
    #>
    Write-Log "Applying System Latency Tweaks..." -Level STEP

    $NetworkPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile"

    try {
        # NetworkThrottlingIndex: FFFFFFFF disables throttling
        Set-ItemProperty -Path $NetworkPath -Name "NetworkThrottlingIndex" -Value 0xFFFFFFFF -Force
        
        # SystemResponsiveness: 0 sets maximum responsiveness for games/media
        Set-ItemProperty -Path $NetworkPath -Name "SystemResponsiveness" -Value 0 -Force
        
        Write-Log "Success: System responsiveness and network throttling optimized." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to apply latency tweaks: $($_.Exception.Message)" -Level WARN
    }
}

function Set-CPUParkingDisable {
    <#
    .SYNOPSIS
        Disables CPU Core Parking to ensure all cores are immediately available.
    #>
    Write-Log "Disabling CPU Core Parking..." -Level STEP
    try {
        & powercfg -setacvalueindex scheme_current sub_processor CPMINCORES 0
        & powercfg -setacvalueindex scheme_current sub_processor CPMAXCORES 100
        & powercfg -setactive scheme_current
        Write-Log "Success: Core Parking disabled." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to disable Core Parking: $($_.Exception.Message)" -Level WARN
    }
}

function Set-CPUPrioritySeparation {
    <#
    .SYNOPSIS
        Optimizes Win32 Priority Separation for maximum foreground responsiveness.
    #>
    Write-Log "Optimizing CPU Priority Separation..." -Level STEP
    try {
        # Value 26 (Hex) / 38 (Dec) for best desktop/gaming performance
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\PriorityControl" -Name "Win32PrioritySeparation" -Value 38 -Force
        Write-Log "Success: CPU Scheduling optimized." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to set Priority Separation: $($_.Exception.Message)" -Level ERROR
    }
}

function Set-AdminContextMenus {
    <#
    .SYNOPSIS
        Adds 'Run CMD' and 'Run PowerShell' (User & Admin) to the directory and background context menus.
    #>
    Write-Log "Adding 'Run CMD' & 'PowerShell' to Context Menus..." -Level STEP

    # 1. Administrator Check
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log "Skipping Context Menus: Administrator rights are required to modify HKEY_CLASSES_ROOT." -Level WARN
        return
    }

    # Define the Registry content with English UI verbs
    $regContent = @"
Windows Registry Editor Version 5.00

# --- CMD Submenu Definitions ---
[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuCmd\shell\open]
"MUIVerb"="Run as User"
"Icon"="cmd.exe"
[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuCmd\shell\open\command]
@="cmd.exe /s /k pushd \"%V\""

[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuCmd\shell\runas]
"MUIVerb"="Run as Administrator"
"Icon"="cmd.exe"
"HasLUAShield"=""
[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuCmd\shell\runas\command]
@="cmd.exe /s /k pushd \"%V\""

# --- PowerShell Submenu Definitions ---
[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuPowerShell\shell\open]
"MUIVerb"="Run as User"
"Icon"="powershell.exe"
[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuPowerShell\shell\open\command]
@="powershell.exe -noexit -command Set-Location '%V'"

[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuPowerShell\shell\runas]
"MUIVerb"="Run as Administrator"
"Icon"="powershell.exe"
"HasLUAShield"=""
[HKEY_CLASSES_ROOT\Directory\ContextMenus\MenuPowerShell\shell\runas\command]
@="powershell.exe -noexit -command Set-Location '%V'"

# --- Main Menu Integration (Folders & Background) ---
[HKEY_CLASSES_ROOT\Directory\shell\01MenuCmd]
"MUIVerb"="Run CMD"
"Icon"="cmd.exe"
"ExtendedSubCommandsKey"="Directory\\ContextMenus\\MenuCmd"

[HKEY_CLASSES_ROOT\Directory\shell\02MenuPowerShell]
"MUIVerb"="Run PowerShell"
"Icon"="powershell.exe"
"ExtendedSubCommandsKey"="Directory\\ContextMenus\\MenuPowerShell"

[HKEY_CLASSES_ROOT\Directory\Background\shell\01MenuCmd]
"MUIVerb"="Run CMD"
"Icon"="cmd.exe"
"ExtendedSubCommandsKey"="Directory\\ContextMenus\\MenuCmd"

[HKEY_CLASSES_ROOT\Directory\Background\shell\02MenuPowerShell]
"MUIVerb"="Run PowerShell"
"Icon"="powershell.exe"
"ExtendedSubCommandsKey"="Directory\\ContextMenus\\MenuPowerShell"
"@

    try {
        # Create a temporary .reg file
        $tempFile = Join-Path $env:TEMP "context_menu_final.reg"
        $regContent | Out-File -FilePath $tempFile -Encoding utf8 -ErrorAction Stop
        Write-Log "Temporary registry file created." -Level INFO

        # Import the registry file using reg.exe
        $importProcess = Start-Process reg.exe -ArgumentList "import `"$tempFile`"" -Wait -PassThru -NoNewWindow
        
        if ($importProcess.ExitCode -eq 0) {
            Write-Log "Registry keys imported successfully." -Level SUCCESS
        } else {
            throw "reg.exe failed with exit code $($importProcess.ExitCode)"
        }

        # Cleanup
        if (Test-Path $tempFile) { Remove-Item $tempFile -ErrorAction SilentlyContinue }

        # Refresh Explorer to apply changes immediately
        Write-Log "Refreshing Windows Explorer shell..." -Level INFO
        $shell = New-Object -ComObject Shell.Application
        $shell.Windows() | ForEach-Object { $_.Refresh() }

        Write-Log "Success: 'Run CMD' and 'Run PowerShell' added to Context Menu." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to update Context Menus: $($_.Exception.Message)" -Level ERROR
    }
}

function Uninstall-OfficeAndOneDrive {
    <#
    .SYNOPSIS
        Uninstalls Microsoft 365 (Office), all Teams variants, and completely purges OneDrive.
    #>
    Write-Log "Radical Removal of Office 365, Teams & OneDrive..." -Level STEP

    try {
        # 1. Remove Office / Microsoft 365
        $officePackages = @(
            "Microsoft.Office.Desktop", 
            "Microsoft.Microsoft365Apps-en-us", 
            "Microsoft.Microsoft365Apps-de-de"
        )
        foreach ($pkg in $officePackages) {
            Write-Log "Checking for $pkg..." -Level INFO
            winget uninstall --id $pkg --silent --accept-source-agreements --ignore-uninstalled | Out-Null
        }

        # 2. Comprehensive Teams Removal
        # This covers: New Teams (Work/School), Classic, and Personal (Free)
        $teamsIds = @(
            "Microsoft.Teams",          # New Teams
            "Microsoft.Teams.Classic",  # Classic Teams
            "Microsoft.Teams.Free"      # Personal Version (Win 11 built-in)
        )
        
        foreach ($id in $teamsIds) {
            Write-Log "Purging Teams variant: $id..." -Level INFO
            winget uninstall --id $id --silent --ignore-uninstalled | Out-Null
        }

        # Remove "Teams Machine-Wide Installer" (Registry-based search is faster than WMI)
        $uninstallKeys = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        $machineInstaller = Get-ItemProperty $uninstallKeys | Where-Object { $_.DisplayName -match "Teams Machine-Wide Installer" }
        
        if ($machineInstaller) {
            Write-Log "Removing Teams Machine-Wide Installer via MSIExec..." -Level INFO
            Start-Process "msiexec.exe" -ArgumentList "/X$($machineInstaller.PSChildName) /qn /norestart" -Wait
        }

        # 3. Radical removal of OneDrive
        Write-Log "Removing OneDrive and cleaning Explorer shell..." -Level WARN
        Stop-Process -Name "OneDrive" -ErrorAction SilentlyContinue
        
        $oneDrivePaths = @(
            "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
            "$env:SystemRoot\System32\OneDriveSetup.exe",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\Update\OneDriveSetup.exe",
            "$(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UninstallString -ErrorAction SilentlyContinue)"
        )

        $uninstallerFound = $false
        foreach ($path in $oneDrivePaths) {
            $cleanPath = $path -replace '"', ''
            
            if (Test-Path $cleanPath) {
                Write-Log "OneDrive Uninstaller found at: $cleanPath. Executing..." -Level INFO
                Start-Process -FilePath $cleanPath -ArgumentList "/uninstall" -Wait
                $uninstallerFound = $true
                break
            }
        }

        if (-not $uninstallerFound) {
            Write-Log "OneDrive Uninstaller not found via standard paths. It might already be removed." -Level INFO
        }

        # 4. Cleanup Registry (Explorer Sidebar)
        $registryPaths = @(
            "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
            "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
        )
        foreach ($path in $registryPaths) {
            if (Test-Path $path) {
                Set-ItemProperty -Path $path -Name "System.IsPinnedToNameSpaceTree" -Value 0 -ErrorAction SilentlyContinue
            }
        }

        # 5. Clean up local folders
        $folders = @(
            "$env:USERPROFILE\OneDrive",
            "$env:LOCALAPPDATA\Microsoft\OneDrive",
            "$env:PROGRAMDATA\Microsoft OneDrive",
            "$env:APPDATA\Microsoft\Teams"
        )
        foreach ($folder in $folders) {
            if (Test-Path $folder) { 
                Remove-Item $folder -Recurse -Force -ErrorAction SilentlyContinue 
            }
        }

        Write-Log "Success: Office 365, all Teams variants, and OneDrive have been purged." -Level SUCCESS
    }
    catch {
        Write-Log "An error occurred during the purge: $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-SystemBloatware {
    <#
    .SYNOPSIS
        Removes all Windows bloatware system-wide and prevents re-installation for new users.
    #>
    Write-Log "Starting Radical System Bloatware Removal..." -Level STEP

    # 1. Extensive list of Appx packages to purge
    $bloatApps = @(
        # Communication
        "Microsoft.Teams", "Microsoft.Teams.Classic", "Microsoft.Teams.Free", 
        "Microsoft.WindowsCommunicationsApps", "Microsoft.SkypeApp", "Microsoft.Messaging",
        
        # Bing & News & Widgets
        "Microsoft.BingNews", "Microsoft.BingWeather", "Microsoft.BingFinance", 
        "Microsoft.BingSports", "Microsoft.BingSearch", "Microsoft.Windows.Widgets.Feeds",
        
        # Media & Entertainment
        "Microsoft.ZuneVideo", "Microsoft.ZuneMusic", "Microsoft.MicrosoftSolitaireCollection", 
        "Clipchamp.Clipchamp", "Microsoft.GamingApp", "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.Xbox.TCUI", "Microsoft.XboxGameOverlay", "Microsoft.XboxIdentityProvider",
        
        # Tools & Helpers
        "Microsoft.GetHelp", "Microsoft.Getstarted", "Microsoft.OneConnect", 
        "Microsoft.People", "Microsoft.YourPhone", "Microsoft.WindowsFeedbackHub", 
        "Microsoft.MicrosoftPowerBIForWindows", "Microsoft.549981C3F5F10", # Cortana
        "Microsoft.WindowsMaps", "Microsoft.Wallet", "Microsoft.MixedReality.Portal",
        "Microsoft.DevHome", "Microsoft.AzureCompute.Edge", "Microsoft.PowerAutomateDesktop",
        "Microsoft.MicrosoftOfficeHub" # Office Stub
    )

    try {
        # 2. Prevent Windows from auto-installing "Consumer Content" (Disney, Candy Crush etc.)
        Write-Log "Disabling automatic consumer content installation..." -Level INFO
        $RegistryPaths = @(
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        )
        foreach ($path in $RegistryPaths) {
            if (!(Test-Path $path)) { New-Item $path -Force | Out-Null }
        }
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SilentInstalledAppsEnabled" -Value 0 -Type DWord

        # 3. Process the main list
        foreach ($appName in $bloatApps) {
            Write-Log "Processing $appName..." -Level INFO
            
            # Remove Provisioned Package (The Template for new users)
            Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -match $appName } | ForEach-Object {
                Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null
                Write-Log "  Removed $appName from Provisioned Packages." -Level SUCCESS
            }

            # Remove Installed Package for All Users
            Get-AppxPackage -Name "*$appName*" -AllUsers -ErrorAction SilentlyContinue | ForEach-Object {
                Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction SilentlyContinue | Out-Null
                Write-Log "  Removed $($_.Name) from all user profiles." -Level SUCCESS
            }
        }

        # 4. Special Purge: Teams Machine-Wide Installer (The "Zombie" Installer)
        Write-Log "Searching for Teams Machine-Wide Installer (MSI)..." -Level WARN
        $TeamsMSI = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Teams Machine-Wide Installer" }
        if ($TeamsMSI) {
            Write-Log "Uninstalling Teams Machine-Wide Installer to prevent re-spawning..." -Level INFO
            $TeamsMSI.Uninstall() | Out-Null
        }

        # 5. Third-Party "Stub" Purge (Placeholders)
        Write-Log "Cleaning third-party placeholders (Disney, Spotify, etc.)..." -Level INFO
        $stubPattern = "Disney|Spotify|TikTok|Instagram|Facebook|LinkedIn|Netflix|PrimeVideo"
        Get-AppxPackage -AllUsers | Where-Object { $_.Name -match $stubPattern } | ForEach-Object {
            Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction SilentlyContinue | Out-Null
        }

        # 6. Apply Policies (Chat, Feeds, etc.)
        $policies = @(
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"; Name = "ChatIcon"; Value = 3 },
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"; Name = "EnableFeeds"; Value = 0 }
        )

        foreach ($p in $policies) {
            if (!(Test-Path $p.Path)) { New-Item $p.Path -Force | Out-Null }
            Set-ItemProperty -Path $p.Path -Name $p.Name -Value $p.Value -Type DWord -ErrorAction SilentlyContinue
        }

        Write-Log "Bloatware removal completed successfully." -Level SUCCESS
    }
    catch {
        Write-Log "An error occurred during radical bloatware removal: $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-SystemAutostarts {
    <# 
    .SYNOPSIS: Cleans HKLM Run keys used by installers.
    #>
    Write-Log "Cleaning global Run keys..." -Level STEP
    $Paths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run")
    $Targets = @("Steam", "EA Desktop", "EpicGamesLauncher", "SteelSeriesGG", "AsusUpdateCheck")

    foreach ($Path in $Paths) {
        if (Test-Path $Path) {
            foreach ($T in $Targets) {
                if (Get-ItemProperty $Path -Name $T -ErrorAction SilentlyContinue) {
                    Remove-ItemProperty $Path -Name $T -Force
                    Write-Log "Removed global autostart: $T" -Level SUCCESS
                }
            }
        }
    }
}

function Set-WindowsGPUTweaks {
    <#
    .SYNOPSIS
        Applies OS-level GPU optimizations.
        Includes HAGS and the universal MPO disable for stability.
    #>
    Write-Log "Applying Universal Windows GPU Tweaks..." -Level STEP

    try {
        # 1. Disable Multi-Plane Overlay (MPO)
        # Universal DWM tweak. Prevents flickering/stuttering in browsers & apps.
        Write-Log "  > Disabling MPO (Multi-Plane Overlay)..." -Level INFO
        $dwmPath = "HKLM:\SOFTWARE\Microsoft\Windows\Dwm"
        if (!(Test-Path $dwmPath)) { New-Item $dwmPath -Force | Out-Null }
        Set-ItemProperty -Path $dwmPath -Name "OverlayTestMode" -Value 5 -Type Dword

        # 2. Enable Hardware Accelerated GPU Scheduling (HAGS)
        Write-Log "  > Enabling HAGS..." -Level INFO
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers" -Name "HwSchMode" -Value 2 -Force

        # 3. Disable Transparency & Game DVR
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\System\GameConfigStore" -Name "GameDVR_Enabled" -Value 0 -Force
        
        Write-Log "Success: Universal GPU optimizations applied." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to apply Windows GPU tweaks: $($_.Exception.Message)" -Level WARN
    }
}

function Run-WindowsUpdateInstallation {
    <#
    .SYNOPSIS
        Searches, downloads, and installs Windows Updates.
        This function blocks execution until all updates are processed.
    #>
    Write-Log "Starting Synchronous Windows Update Process..." -Level STEP

    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

        # 1. Search for missing updates
        Write-Log "Searching for applicable updates (this may take a few minutes)..." -Level INFO
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

        if ($SearchResult.Updates.Count -eq 0) {
            Write-Log "System is up to date. No updates found." -Level SUCCESS
            return
        }

        Write-Log "Found $($SearchResult.Updates.Count) updates. Starting Download..." -Level INFO

        # 2. Download Updates
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $SearchResult.Updates
        $Downloader.Download()

        # 3. Install Updates
        Write-Log "Download complete. Starting Installation..." -Level INFO
        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $SearchResult.Updates
        $InstallationResult = $Installer.Install()

        # 4. Result Check
        if ($InstallationResult.ResultCode -eq 2) {
            Write-Log "All updates installed successfully." -Level SUCCESS
        } else {
            Write-Log "Some updates could not be installed. ResultCode: $($InstallationResult.ResultCode)" -Level WARN
        }

        if ($InstallationResult.RebootRequired) {
            Write-Log "CRITICAL: A reboot is required to finish the update process." -Level WARN
        }
    }
    catch {
        Write-Log "An error occurred during Windows Update: $($_.Exception.Message)" -Level ERROR
    }
}

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

$UtilsPath = Join-Path $PSScriptRoot "Utils.ps1"
if (Test-Path $UtilsPath) { . $UtilsPath } else { Throw "Critical Error: Utils.ps1 not found!" }

Write-Log "Starting System-Level Setup" -Level STEP

Assert-Admin

if ($ForceStep -gt 0) { 
    Write-Log "Manual Override: Starting at Step $ForceStep" -Level WARN 
}

# 1. Prepare Windows
if (Confirm-StepExecution "Prepare Windows" 1 $StopAfterStep) {
    Disable-AutomaticDriverInstallation
    Set-DeviceManagerFullTransparency
    Clear-WindowsUpdateCache
}

# 2. System Hardening & Core Paths
if (Confirm-StepExecution "System Hardening & Paths" 2 $StopAfterStep) {
    Set-LegacyPasswordLogin
    Set-UACLevel -Level 4
    Set-PrivacyHardening
    Set-TimeSync -SourceIndex 1
    Set-FirewallHardening
    Enable-NetworkDiscovery
    Set-NetworkPrivate
    Set-HighPerformanceMode
}

# 3. Windows Cleanup (Bloatware Removal)
if (Confirm-StepExecution "Windows Cleanup" 3 $StopAfterStep) {
    Uninstall-OfficeAndOneDrive
    Remove-SystemBloatware
    Disable-TelemetryCompletely
}

# 4. Final System Optimization & Performance
if (Confirm-StepExecution "Final System Optimization" 4 $StopAfterStep) {
    # CPU & Latency Tweaks
    Set-CPUParkingDisable
    Set-CPUPrioritySeparation
    Set-CPULatencyTweaks
    
    # Service Cleanup
    Remove-SystemAutostarts
    
    # GPU Finalization
    Set-WindowsGPUTweaks
}

# 5. Windows Updates
if (Confirm-StepExecution "Windows Updates" 5 $StopAfterStep) {
    Run-WindowsUpdateInstallation
}

Remove-ProgressFile

Write-Host "`n"
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Log "CORE SETUP COMPLETED" -Level SUCCESS
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Log "A system reboot is REQUIRED to finalize all changes." -Level WARN

$cancelled = Wait-ForKeyOrTimeout -Timeout 30 -Message "Rebooting in"

if ($cancelled) {
    Write-Log "Reboot cancelled by user. Review logs above." -Level ERROR
} else {
    Write-Log "Initiating system restart..." -Level INFO
    Restart-Computer -Force
}