# 🔍 laptop-health-check
> PowerShell toolkit to audit, verify, and report hardware component originality in refurbished laptops — supports Dell, HP, and Lenovo.
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://docs.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
---
## 🤔 Why?
Buying a refurbished laptop? The seller says "all original parts" — but how do you actually verify that? This script audits every major hardware component, cross-references manufacturers against known OEM supplier databases, and flags anything suspicious.
**Common things refurbishers swap out:**
- 🔋 Battery (cheaper third-party replacements)
- 💾 RAM (lower-spec or mismatched modules)
- 💿 SSD/HDD (used drives from other machines)
- 📡 WiFi card (non-whitelisted replacements)
- 🖥️ Display panel (lower quality aftermarket screens)
- 🔌 Charger (generic adapters without OEM handshake)
This tool catches all of that.
---
## ✨ Features
- **Multi-brand support** — Dell, HP, and Lenovo with brand-specific OEM vendor databases
- **Auto-detection** — automatically identifies your laptop brand, or specify manually
- **12 hardware checks** — CPU, RAM, storage, battery, display, WiFi, audio, webcam, charger, TPM, and more
- **S.M.A.R.T. analysis** — reads power-on hours and SSD wear level to detect reused drives
- **Battery health** — reports design vs actual capacity, health %, and charge cycle count
- **Event log scan** — checks Windows logs for AC adapter / charger warnings
- **Color-coded output** — ✅ green for verified, ⚠️ yellow for flagged items
- **Multiple export formats** — text, HTML dashboard, and JSON
- **Zero dependencies** — pure PowerShell, no third-party tools needed
---
## 🚀 Quick Start
### 1. Download
```powershell
# Clone the repo
git clone https://github.com/YOUR_USERNAME/laptop-health-check.git
cd laptop-health-check
# Or just download the script directly
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/YOUR_USERNAME/laptop-health-check/main/Laptop-ComponentCheck.ps1" -OutFile "Laptop-ComponentCheck.ps1"
```
### 2. Run
```powershell
# Auto-detect brand and run all checks
powershell -ExecutionPolicy Bypass -File .\Laptop-ComponentCheck.ps1 -Brand Dell -ExportHtml
# Specify brand explicitly
.\Laptop-ComponentCheck.ps1 -Brand Dell
.\Laptop-ComponentCheck.ps1 -Brand HP
.\Laptop-ComponentCheck.ps1 -Brand Lenovo
```
> [!TIP]
> **Run as Administrator** for full results. Right-click PowerShell → "Run as Administrator".  
