# Windows 11 Tweaks

> ⚠️ **WORK IN PROGRESS (WIP)** — This repository is a subset of a larger project. It currently handles initial system and user-level optimizations. Features like Driver Management, and Automated Software Deployment are under development.

---

## 🛑 Hardware Disclaimer & Personalization
**Important:** These scripts are currently **tailored specifically for my personal hardware and workflow.** * **Limited Scope:** To prevent compatibility issues for other users, hardware-specific modules (such as drivers or low-level firmware tweaks) that work only for my specific setup are **currently omitted** from this public version.
* **Review the code** before running it on different hardware to ensure compatibility with your specific environment.

---

## Prerequisites: Execution Policy

PowerShell restricts script execution by default. You must grant permission to run these local scripts before starting.

Open PowerShell and run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

This allows local scripts to run while requiring digital signatures for scripts downloaded from the internet.

---

## Permissions & Execution Context

> **Crucial:** Windows manages "System" and "User" settings in different registry hives. You must run the scripts in their designated contexts:

### 1. System-Level (`SystemCleanup.ps1`)
- **Context:** Administrator *(Right-click PowerShell → Run as Administrator)*
- **Target:** `HKLM` (System-wide), Services, Global Bloatware

### 2. User-Level (`UserCleanup.ps1`)
- **Context:** Standard User *(Run normally, **DO NOT** use Admin)*
- **Target:** `HKCU` (Current Profile Hive NTUSER.DAT), Taskbar, Personal Themes, Explorer
- **Note:** Running as Admin will apply changes to the Admin profile, not your daily account!

---

## Step Index & Parameters

Both scripts use a modular engine. Use the `-ForceStep` and `-StopAfterStep` parameters to control the flow.

### `SystemCleanup.ps1` — System Modules

| Step # | Module Name | Description |
|--------|-------------|-------------|
| 1 | Prepare Windows | Disables automatic driver installation, enables Device Manager transparency, clears Windows Update cache. |
| 2 | System Hardening & Core Paths | Legacy login, UAC level, privacy hardening, time sync, firewall, network config, high performance mode. |
| 3 | Windows Cleanup | Uninstalls Office & OneDrive, removes system bloatware, disables telemetry completely. |
| 4 | Final System Optimization | CPU parking/priority/latency tweaks, service autostart cleanup, GPU tweaks. |
| 5 | Windows Updates | Runs Windows Update installation. |

### `UserSetup.ps1` — User Modules

| Step # | Module Name | Description |
|--------|-------------|-------------|
| 1 | Visuals & Themes | Explorer tweaks, Classic Context Menu, color personalization. |
| 2 | Taskbar & Interface | Taskbar config, UI optimization, removes Bing from Start Menu. |
| 3 | Input & Language | System language, keyboard layout, disables Mouse Acceleration (1:1). |
| 4 | Start Menu & Bloatware | Optimizes Start Menu, removes user-level bloatware via `winget`. |
| 5 | Applications & Autostart | Removes unwanted user autostarts (Edge, Spotify, OneDrive, etc.). |
| 6 | Finalization | Restarts Windows Explorer to apply all changes. |

### Example Command

```powershell
# Only run Taskbar and Input settings
.\UserSetup.ps1 -ForceStep 2 -StopAfterStep 3
```

---

## Workflow

1. **System Optimization:** Open PowerShell as Administrator and run `.\SystemCleanup.ps1`
2. **User Personalization:** Open a standard PowerShell window and run `.\UserCleanup.ps1`
