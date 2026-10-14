<#
.SYNOPSIS
    Laptop Component Originality Checker — Dell, HP, Lenovo
.DESCRIPTION
    Audits hardware components in a refurbished laptop and generates a
    detailed report to verify whether parts are original or swapped.
    Supports Dell, HP, and Lenovo with brand-specific OEM validation.
.PARAMETER Brand
    The laptop manufacturer: Dell, HP, Lenovo, or Auto (default).
    Auto will detect the brand from system info.
.PARAMETER ExportHtml
    Generate an HTML report alongside the text report.
.PARAMETER ExportJson
    Export raw data as a JSON file for programmatic use.
.PARAMETER OutputPath
    Base path for report files (without extension).
    Defaults to .\LaptopComponentReport
.PARAMETER SkipBatteryReport
    Skip generating the powercfg battery report.
.EXAMPLE
    .\Laptop-ComponentCheck.ps1
    .\Laptop-ComponentCheck.ps1 -Brand Dell -ExportHtml
    .\Laptop-ComponentCheck.ps1 -Brand HP -ExportHtml -ExportJson
    .\Laptop-ComponentCheck.ps1 -Brand Lenovo -OutputPath "C:\Reports\MyLaptop"
.NOTES
    Run as Administrator for full hardware access.
    Author : Antigravity
    Date   : 2026-10-14
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet("Auto", "Dell", "HP", "Lenovo")]
    [string]$Brand = "Auto",

    [switch]$ExportHtml,
    [switch]$ExportJson,

    [string]$OutputPath = "$PSScriptRoot\LaptopComponentReport",

    [switch]$SkipBatteryReport
)

# ═══════════════════════════════════════════════════════════════════════════
#  BRAND-SPECIFIC CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════

$BrandProfiles = @{

    Dell = @{
        DisplayName     = "Dell"
        MatchPatterns   = @("*Dell*")
        SupportUrl      = "https://www.dell.com/support/home/en-us/product-support/servicetag/{0}/overview"
        SerialLabel     = "Service Tag"
        # Known OEM RAM suppliers
        RamVendors      = @("Samsung", "SK Hynix", "Hynix", "Micron", "Kingston",
                            "Ramaxel", "Smart Modular", "Nanya", "Elpida")
        # Known OEM battery manufacturers
        BatteryVendors  = @("Dell", "Samsung SDI", "LGC", "Simplo", "SMP",
                            "BYD", "Celxpert", "Sunwoda")
        # Known WiFi card vendors (Dell uses BIOS whitelist on many models)
        WifiVendors     = @("Intel", "Qualcomm", "Killer", "Dell Wireless",
                            "Broadcom", "Rivet Networks", "MediaTek")
        WifiWhitelist   = $true   # Dell enforces BIOS-level WiFi whitelist
        # Known display panel manufacturers
        PanelVendors    = @("AUO", "BOE", "LGD", "CMN", "SDC", "SHP", "IVO",
                            "AU Optronics", "LG Display", "Sharp", "Innolux")
        # Known SSD/HDD OEM suppliers
        StorageVendors  = @("Samsung", "SK Hynix", "Hynix", "Micron", "Toshiba",
                            "KIOXIA", "Western Digital", "WDC", "Intel", "Lite-On",
                            "SanDisk", "Seagate", "LITEON", "KXG", "SSDPEM", "MZVL")
        BIOSVendor      = "Dell Inc."
        Notes           = @(
            "Dell laptops often have a BIOS WiFi whitelist — non-approved cards may not boot."
            "Service Tag can be verified at dell.com/support for original factory configuration."
            "Dell batteries use Dell-specific connectors and firmware handshake."
        )
    }

    HP = @{
        DisplayName     = "HP"
        MatchPatterns   = @("*HP*", "*Hewlett*Packard*", "*Hewlett-Packard*")
        SupportUrl      = "https://support.hp.com/us-en/product/details/{0}"
        SerialLabel     = "Serial Number"
        RamVendors      = @("Samsung", "SK Hynix", "Hynix", "Micron", "Kingston",
                            "Ramaxel", "Smart Modular", "Nanya", "Elpida")
        BatteryVendors  = @("HP", "Hewlett-Packard", "Samsung SDI", "LGC", "Simplo",
                            "SMP", "BYD", "Celxpert", "Sunwoda", "Dynapack")
        WifiVendors     = @("Intel", "Qualcomm", "Broadcom", "Realtek", "MediaTek",
                            "Ralink")
        WifiWhitelist   = $true   # Many HP models enforce BIOS WiFi whitelist
        PanelVendors    = @("AUO", "BOE", "LGD", "CMN", "SDC", "SHP", "IVO",
                            "AU Optronics", "LG Display", "Sharp", "Innolux")
        StorageVendors  = @("Samsung", "SK Hynix", "Hynix", "Micron", "Toshiba",
                            "KIOXIA", "Western Digital", "WDC", "Intel", "Lite-On",
                            "SanDisk", "Seagate", "LITEON", "KXG", "SSDPEM", "MZVL")
        BIOSVendor      = "HP"
        Notes           = @(
            "HP laptops often have a BIOS WiFi whitelist — non-approved cards may prevent boot."
            "Use HP Support Assistant or hp.com/support to verify original configuration."
            "HP product number (P/N) is different from serial — both are useful for verification."
        )
    }

    Lenovo = @{
        DisplayName     = "Lenovo"
        MatchPatterns   = @("*Lenovo*", "*LENOVO*", "*ThinkPad*", "*IdeaPad*")
        SupportUrl      = "https://pcsupport.lenovo.com/us/en/products/{0}"
        SerialLabel     = "Serial Number"
        RamVendors      = @("Samsung", "SK Hynix", "Hynix", "Micron", "Kingston",
                            "Ramaxel", "Smart Modular", "Nanya", "Elpida")
        BatteryVendors  = @("Lenovo", "LGC", "Samsung SDI", "Simplo", "SMP",
                            "BYD", "Celxpert", "Sunwoda", "SanYo", "Panasonic",
                            "CPT", "Dynapack")
        WifiVendors     = @("Intel", "Qualcomm", "Broadcom", "Realtek", "MediaTek",
                            "Fibocom", "Ralink")
        WifiWhitelist   = $false  # Newer ThinkPads removed whitelist; older ones had it
        PanelVendors    = @("AUO", "BOE", "LGD", "CMN", "SDC", "SHP", "IVO",
                            "AU Optronics", "LG Display", "Sharp", "Innolux", "CSO")
        StorageVendors  = @("Samsung", "SK Hynix", "Hynix", "Micron", "Toshiba",
                            "KIOXIA", "Western Digital", "WDC", "Intel", "Lite-On",
                            "SanDisk", "Seagate", "LITEON", "Union Memory",
                            "UMIS", "KXG", "SSDPEM", "MZVL")
        BIOSVendor      = "LENOVO"
        Notes           = @(
            "Older ThinkPads had a BIOS WiFi whitelist; newer models (2019+) removed it."
            "Lenovo Vantage or pcsupport.lenovo.com can verify original specs by serial."
            "ThinkPad batteries have an embedded controller chip that reports cycle count."
        )
    }
}

# ═══════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

function Write-Banner {
    param([string]$Title)
    $width = 64
    $pad = [math]::Max(0, [math]::Floor(($width - $Title.Length - 2) / 2))
    $line = "═" * $width
    Write-Host "`n$line" -ForegroundColor DarkCyan
    Write-Host (" " * $pad + "  $Title") -ForegroundColor Cyan
    Write-Host "$line" -ForegroundColor DarkCyan
}

function Write-Section {
    param([string]$Title)
    $sep = "─" * 60
    Write-Host "`n$sep" -ForegroundColor DarkCyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "$sep" -ForegroundColor DarkCyan
}

function Write-Field {
    param([string]$Label, [string]$Value, [ConsoleColor]$Color = 'White')
    Write-Host ("  {0,-30} : " -f $Label) -NoNewline -ForegroundColor Gray
    Write-Host $Value -ForegroundColor $Color
}

function Write-StatusField {
    param([string]$Label, [string]$Value, [ValidateSet('OK','Warn','Info')]$Status = 'Info')
    $icon  = switch ($Status) { 'OK' { "✓" } 'Warn' { "⚠" } 'Info' { "ℹ" } }
    $color = switch ($Status) { 'OK' { 'Green' } 'Warn' { 'Yellow' } 'Info' { 'DarkGray' } }
    Write-Field -Label "$icon $Label" -Value $Value -Color $color
}

function Get-FriendlySize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N1} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N0} KB" -f ($Bytes / 1KB)
}

function Test-VendorMatch {
    param(
        [string]$ActualValue,
        [string[]]$KnownVendors
    )
    if (-not $ActualValue) { return $false }
    foreach ($vendor in $KnownVendors) {
        if ($ActualValue -like "*$vendor*") { return $true }
    }
    return $false
}

function Decode-WmiString {
    param([uint16[]]$Bytes)
    if (-not $Bytes) { return "" }
    ($Bytes | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
}

# ═══════════════════════════════════════════════════════════════════════════
#  INITIALIZATION
# ═══════════════════════════════════════════════════════════════════════════

$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$report   = [ordered]@{}
$warnings = [System.Collections.Generic.List[string]]::new()

# ── Detect / validate brand ─────────────────────────────────────────────────

$cs   = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$bb   = Get-CimInstance Win32_BaseBoard

$detectedManufacturer = $cs.Manufacturer

if ($Brand -eq "Auto") {
    $Brand = $null
    foreach ($brandName in $BrandProfiles.Keys) {
        $profile = $BrandProfiles[$brandName]
        foreach ($pattern in $profile.MatchPatterns) {
            if ($detectedManufacturer -like $pattern) {
                $Brand = $brandName
                break
            }
        }
        if ($Brand) { break }
    }
    if (-not $Brand) {
        Write-Host "`n⚠  Could not auto-detect brand from manufacturer: '$detectedManufacturer'" -ForegroundColor Yellow
        Write-Host "   Supported brands: Dell, HP, Lenovo" -ForegroundColor Yellow
        Write-Host "   Falling back to generic checks.`n" -ForegroundColor Yellow
        $Brand = "Generic"
    } else {
        Write-Host "`n✓  Auto-detected brand: $Brand (from '$detectedManufacturer')" -ForegroundColor Green
    }
}

# Build effective profile (generic fallback if brand not in list)
if ($BrandProfiles.ContainsKey($Brand)) {
    $profile = $BrandProfiles[$Brand]
} else {
    $profile = @{
        DisplayName    = "Generic"
        MatchPatterns  = @()
        SupportUrl     = ""
        SerialLabel    = "Serial Number"
        RamVendors     = @("Samsung", "SK Hynix", "Micron", "Kingston", "Ramaxel", "Nanya")
        BatteryVendors = @()
        WifiVendors    = @("Intel", "Qualcomm", "Broadcom", "Realtek", "MediaTek")
        WifiWhitelist  = $false
        PanelVendors   = @("AUO", "BOE", "LGD", "CMN", "SDC", "SHP", "IVO")
        StorageVendors = @("Samsung", "SK Hynix", "Micron", "Toshiba", "KIOXIA",
                           "Western Digital", "WDC", "Intel", "SanDisk", "Seagate")
        BIOSVendor     = ""
        Notes          = @()
    }
}

# ═══════════════════════════════════════════════════════════════════════════
#  HEADER
# ═══════════════════════════════════════════════════════════════════════════

Write-Banner "$($profile.DisplayName) LAPTOP COMPONENT ORIGINALITY CHECKER"

if (-not $isAdmin) {
    Write-Host "`n  ⚠  Running without Administrator privileges. Some checks will be limited." -ForegroundColor Yellow
    Write-Host "     Recommend: Right-click PowerShell → 'Run as Administrator'`n" -ForegroundColor DarkYellow
}

# ═══════════════════════════════════════════════════════════════════════════
#  1. SYSTEM IDENTITY
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "1. SYSTEM IDENTITY"

$serialNumber = $bios.SerialNumber
$model        = $cs.Model
$biosVersion  = $bios.SMBIOSBIOSVersion
$biosDate     = $bios.ReleaseDate

$report['System'] = [ordered]@{
    Brand          = $profile.DisplayName
    Manufacturer   = $detectedManufacturer
    Model          = $model
    ($profile.SerialLabel) = $serialNumber
    BIOSVersion    = $biosVersion
    BIOSDate       = $biosDate?.ToString("yyyy-MM-dd")
    Motherboard    = "$($bb.Manufacturer) / $($bb.Product)"
    MoboSerial     = $bb.SerialNumber
    SystemSKU      = $cs.SystemSKUNumber
}

Write-Field "Manufacturer" $detectedManufacturer
Write-Field "Model" $model
Write-Field $profile.SerialLabel $serialNumber

# Brand validation
$brandMatch = $false
foreach ($pattern in $profile.MatchPatterns) {
    if ($detectedManufacturer -like $pattern) { $brandMatch = $true; break }
}

if ($Brand -ne "Generic" -and -not $brandMatch) {
    Write-StatusField "Brand Mismatch" "Expected $($profile.DisplayName) but detected '$detectedManufacturer'" -Status Warn
    $warnings.Add("System manufacturer '$detectedManufacturer' does not match expected brand '$($profile.DisplayName)'.")
} elseif ($Brand -ne "Generic") {
    Write-StatusField "Brand Verified" "$($profile.DisplayName) system confirmed" -Status OK
}

Write-Field "BIOS Version" $biosVersion
Write-Field "BIOS Date" ($biosDate?.ToString("yyyy-MM-dd"))
Write-Field "Motherboard" "$($bb.Manufacturer) / $($bb.Product)"
Write-Field "Motherboard Serial" $bb.SerialNumber
if ($cs.SystemSKUNumber) { Write-Field "System SKU" $cs.SystemSKUNumber }

# Brand-specific support link
if ($profile.SupportUrl) {
    $supportUrl = $profile.SupportUrl -f $serialNumber
    Write-Host "`n  🔗 " -NoNewline
    Write-Host "Verify original config: " -NoNewline -ForegroundColor Gray
    Write-Host $supportUrl -ForegroundColor Blue
} else {
    $supportUrl = ""
}

# HP-specific: also grab product number
if ($Brand -eq "HP") {
    $hpProductNum = $cs.SystemSKUNumber
    if ($hpProductNum) {
        Write-Field "HP Product Number (P/N)" $hpProductNum
        $hpLookupUrl = "https://support.hp.com/us-en/product/details/$hpProductNum"
        Write-Host "  🔗 " -NoNewline
        Write-Host "HP P/N lookup: " -NoNewline -ForegroundColor Gray
        Write-Host $hpLookupUrl -ForegroundColor Blue
    }
}

# Lenovo-specific: machine type
if ($Brand -eq "Lenovo") {
    # Lenovo model string often includes machine type like "20Y3CTO1WW"
    $lenovoMachineType = ($model -split '\s')[0]
    if ($lenovoMachineType) {
        Write-Field "Lenovo Machine Type" $lenovoMachineType
    }
}

# ═══════════════════════════════════════════════════════════════════════════
#  2. PROCESSOR (CPU)
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "2. PROCESSOR (CPU)"

$cpus = Get-CimInstance Win32_Processor
foreach ($cpu in $cpus) {
    $cpuArch = switch ($cpu.Architecture) {
        0 {"x86"} 5 {"ARM"} 9 {"x64"} 12 {"ARM64"} default {"Unknown ($($cpu.Architecture))"}
    }

    $report['CPU'] = [ordered]@{
        Name         = $cpu.Name.Trim()
        Cores        = $cpu.NumberOfCores
        Threads      = $cpu.NumberOfLogicalProcessors
        MaxClockMHz  = $cpu.MaxClockSpeed
        Socket       = $cpu.SocketDesignation
        Architecture = $cpuArch
    }

    Write-Field "Name" $cpu.Name.Trim()
    Write-Field "Cores / Threads" "$($cpu.NumberOfCores) / $($cpu.NumberOfLogicalProcessors)"
    Write-Field "Max Clock Speed" "$($cpu.MaxClockSpeed) MHz"
    Write-Field "Architecture" $cpuArch
    Write-Field "Socket" $cpu.SocketDesignation

    # Note on replaceability
    $solderedIndicators = @("BGA", "FP", "U3E1")
    $isSoldered = $solderedIndicators | Where-Object { $cpu.SocketDesignation -like "*$_*" }
    if ($isSoldered -or $cpu.UpgradeMethod -in @(1, 2, 6)) {
        Write-StatusField "CPU Type" "Likely soldered — cannot be swapped" -Status OK
    } else {
        Write-StatusField "CPU Type" "Socketed — CPU could have been replaced (verify against original spec)" -Status Info
    }
}

# ═══════════════════════════════════════════════════════════════════════════
#  3. MEMORY (RAM)
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "3. MEMORY (RAM)"

$ramModules = Get-CimInstance Win32_PhysicalMemory
$totalRAM   = ($ramModules | Measure-Object -Property Capacity -Sum).Sum

Write-Field "Total Installed" (Get-FriendlySize $totalRAM)
Write-Field "Module Count" $ramModules.Count

$report['RAM'] = [System.Collections.Generic.List[object]]::new()

$moduleIndex = 0
foreach ($ram in $ramModules) {
    $moduleIndex++

    $mfr      = if ($ram.Manufacturer) { $ram.Manufacturer.Trim() } else { "Unknown" }
    $partNum  = if ($ram.PartNumber)    { $ram.PartNumber.Trim() }    else { "Unknown" }
    $capacity = Get-FriendlySize $ram.Capacity
    $speed    = "$($ram.Speed) MHz"
    $type     = switch ($ram.SMBIOSMemoryType) {
        20 {"DDR"} 21 {"DDR2"} 22 {"DDR2"} 24 {"DDR3"} 26 {"DDR4"} 30 {"LPDDR4"} 34 {"DDR5"} 35 {"LPDDR5"}
        default {
            if ($ram.MemoryType -eq 0 -and $ram.Speed -ge 4800) { "DDR5 (inferred)" }
            elseif ($ram.MemoryType -eq 0 -and $ram.Speed -ge 2133) { "DDR4 (inferred)" }
            else { "Type $($ram.SMBIOSMemoryType)" }
        }
    }

    $formFactor = switch ($ram.FormFactor) {
        8 {"DIMM"} 12 {"SO-DIMM"} 0 {"Unknown (likely soldered)"}
        default { "Form $($ram.FormFactor)" }
    }

    Write-Host ""
    Write-Field "  Module $moduleIndex" "$capacity  |  $type  |  $speed  |  $formFactor"
    Write-Field "  Manufacturer" $mfr
    Write-Field "  Part Number" $partNum
    Write-Field "  Slot" "$($ram.BankLabel) / $($ram.DeviceLocator)"

    $isKnownMfr = Test-VendorMatch -ActualValue $mfr -KnownVendors $profile.RamVendors
    if (-not $isKnownMfr -and $mfr -ne "Unknown") {
        Write-StatusField "  RAM Vendor" "'$mfr' is NOT a common $($profile.DisplayName) OEM RAM supplier" -Status Warn
        $warnings.Add("RAM Module $moduleIndex: Manufacturer '$mfr' is uncommon for $($profile.DisplayName) OEM.")
    } elseif ($mfr -ne "Unknown") {
        Write-StatusField "  RAM Vendor" "Known $($profile.DisplayName) OEM supplier" -Status OK
    }

    $report['RAM'].Add([ordered]@{
        Module       = $moduleIndex
        Capacity     = $capacity
        Type         = $type
        Speed        = $speed
        FormFactor   = $formFactor
        Manufacturer = $mfr
        PartNumber   = $partNum
        Slot         = "$($ram.BankLabel) / $($ram.DeviceLocator)"
    })
}

# ═══════════════════════════════════════════════════════════════════════════
#  4. STORAGE DRIVES
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "4. STORAGE DRIVES"

$disks = Get-CimInstance Win32_DiskDrive | Where-Object { $_.MediaType -notlike "*Removable*" }
$report['Storage'] = [System.Collections.Generic.List[object]]::new()

$driveIndex = 0
foreach ($disk in $disks) {
    $driveIndex++
    $sizeStr = Get-FriendlySize $disk.Size

    Write-Host ""
    Write-Field "  Drive $driveIndex" $disk.Model
    Write-Field "  Serial Number" $disk.SerialNumber?.Trim()
    Write-Field "  Capacity" $sizeStr
    Write-Field "  Interface" $disk.InterfaceType
    Write-Field "  Media Type" $disk.MediaType
    Write-Field "  Firmware Rev" $disk.FirmwareRevision

    $driveData = [ordered]@{
        Model        = $disk.Model
        SerialNumber = $disk.SerialNumber?.Trim()
        Capacity     = $sizeStr
        Interface    = $disk.InterfaceType
        MediaType    = $disk.MediaType
        FirmwareRev  = $disk.FirmwareRevision
    }

    # Check vendor
    $isKnownStorage = Test-VendorMatch -ActualValue $disk.Model -KnownVendors $profile.StorageVendors
    if ($isKnownStorage) {
        Write-StatusField "  Storage Vendor" "Known OEM drive manufacturer" -Status OK
    } elseif ($disk.Model) {
        Write-StatusField "  Storage Vendor" "'$($disk.Model)' — verify with $($profile.DisplayName) original config" -Status Warn
        $warnings.Add("Drive $driveIndex: '$($disk.Model)' — verify against original configuration.")
    }

    # S.M.A.R.T. / power-on hours
    if ($isAdmin) {
        try {
            $physDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue
            foreach ($pd in $physDisks) {
                if ($disk.Model -like "*$($pd.FriendlyName)*" -or $pd.FriendlyName -like "*$($disk.Model.Trim())*" -or $pd.DeviceId -eq ($driveIndex - 1)) {
                    Write-Field "  Health Status" $pd.HealthStatus
                    $driveData['HealthStatus'] = $pd.HealthStatus

                    $rel = $pd | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
                    if ($rel) {
                        if ($rel.PowerOnHours) {
                            $poh = $rel.PowerOnHours
                            $pohDays = [math]::Round($poh / 24)
                            Write-Field "  Power-On Hours" "$poh hrs (~$pohDays days)"
                            $driveData['PowerOnHours'] = $poh

                            if ($poh -gt 15000) {
                                Write-StatusField "  Drive Usage" "Very high ($poh hrs) — likely reused from another machine" -Status Warn
                                $warnings.Add("Drive $driveIndex ($($disk.Model)): $poh power-on hours — likely not original.")
                            } elseif ($poh -gt 8000) {
                                Write-StatusField "  Drive Usage" "Moderate ($poh hrs) — verify against expected refurb age" -Status Warn
                            } else {
                                Write-StatusField "  Drive Usage" "$poh hrs — looks reasonable" -Status OK
                            }
                        }
                        if ($rel.Temperature) {
                            Write-Field "  Temperature" "$($rel.Temperature)°C"
                            $driveData['Temperature'] = $rel.Temperature
                        }
                        if ($null -ne $rel.Wear) {
                            Write-Field "  SSD Wear Level" "$($rel.Wear)%"
                            $driveData['WearLevel'] = $rel.Wear
                            if ($rel.Wear -gt 50) {
                                Write-StatusField "  Wear" "Above 50% — SSD has significant wear" -Status Warn
                                $warnings.Add("Drive $driveIndex: SSD wear at $($rel.Wear)%.")
                            }
                        }
                        if ($rel.ReadErrorsTotal)  { Write-Field "  Read Errors" $rel.ReadErrorsTotal }
                        if ($rel.WriteErrorsTotal) { Write-Field "  Write Errors" $rel.WriteErrorsTotal }
                    }
                    break
                }
            }
        } catch {
            Write-Host "    ℹ  Could not retrieve S.M.A.R.T. data for this drive" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "    ℹ  Run as Admin to get power-on hours and health data" -ForegroundColor DarkGray
    }

    $report['Storage'].Add($driveData)
}

# ═══════════════════════════════════════════════════════════════════════════
#  5. BATTERY
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "5. BATTERY"

$batteries = Get-CimInstance Win32_Battery
$batteryDetails = Get-CimInstance -Namespace root\WMI -ClassName BatteryStaticData -ErrorAction SilentlyContinue
$batteryStatus  = Get-CimInstance -Namespace root\WMI -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue
$batteryCycle   = Get-CimInstance -Namespace root\WMI -ClassName BatteryCycleCount -ErrorAction SilentlyContinue

$report['Battery'] = [System.Collections.Generic.List[object]]::new()

if ($batteries) {
    foreach ($bat in $batteries) {
        $chemistry = switch ($bat.Chemistry) {
            1 {"Other"} 2 {"Unknown"} 3 {"Lead Acid"} 4 {"Nickel Cadmium"}
            5 {"Nickel Metal Hydride"} 6 {"Lithium-ion"} 7 {"Zinc Air"}
            8 {"Lithium Polymer"} default {"Code $($bat.Chemistry)"}
        }

        $batData = [ordered]@{
            Name      = $bat.Name
            Status    = $bat.Status
            Chemistry = $chemistry
        }

        Write-Field "Name" $bat.Name
        Write-Field "Status" $bat.Status
        Write-Field "Chemistry" $chemistry
        if ($bat.DesignVoltage) {
            Write-Field "Design Voltage" "$($bat.DesignVoltage) mV"
            $batData['DesignVoltage'] = "$($bat.DesignVoltage) mV"
        }
    }

    # Detailed WMI battery info
    if ($batteryDetails) {
        foreach ($bd in $batteryDetails) {
            $mfr = $bd.ManufactureName
            Write-Field "Manufacturer" $mfr

            $isOemBattery = Test-VendorMatch -ActualValue $mfr -KnownVendors $profile.BatteryVendors
            if ($isOemBattery) {
                Write-StatusField "Battery Origin" "'$mfr' is a known $($profile.DisplayName) battery supplier" -Status OK
            } elseif ($mfr) {
                Write-StatusField "Battery Origin" "'$mfr' is NOT a known $($profile.DisplayName) battery supplier" -Status Warn
                $warnings.Add("Battery manufacturer '$mfr' is not a known $($profile.DisplayName) OEM supplier.")
            }

            if ($bd.UniqueID)       { Write-Field "Unique ID" $bd.UniqueID }
            if ($bd.DeviceName)     { Write-Field "Device Name" $bd.DeviceName }
            if ($bd.DesignedCapacity) {
                $designWh = [math]::Round($bd.DesignedCapacity / 1000, 1)
                Write-Field "Design Capacity" "$($bd.DesignedCapacity) mWh (~${designWh} Wh)"
            }
        }
    }

    # Full charge capacity + health %
    if ($batteryStatus) {
        foreach ($bs in $batteryStatus) {
            if ($bs.FullChargedCapacity) {
                $fullWh = [math]::Round($bs.FullChargedCapacity / 1000, 1)
                Write-Field "Full Charge Capacity" "$($bs.FullChargedCapacity) mWh (~${fullWh} Wh)"

                if ($batteryDetails -and $batteryDetails[0].DesignedCapacity -gt 0) {
                    $healthPct = [math]::Round(($bs.FullChargedCapacity / $batteryDetails[0].DesignedCapacity) * 100, 1)
                    if ($healthPct -ge 80) {
                        Write-StatusField "Battery Health" "$healthPct% of design capacity — good condition" -Status OK
                    } elseif ($healthPct -ge 50) {
                        Write-StatusField "Battery Health" "$healthPct% of design capacity — moderate wear" -Status Warn
                        $warnings.Add("Battery health at $healthPct% — moderate degradation.")
                    } else {
                        Write-StatusField "Battery Health" "$healthPct% of design capacity — significant degradation!" -Status Warn
                        $warnings.Add("Battery health critically low at $healthPct%!")
                    }
                }
            }
        }
    }

    # Cycle count
    if ($batteryCycle) {
        foreach ($bc in $batteryCycle) {
            if ($bc.CycleCount) {
                Write-Field "Cycle Count" $bc.CycleCount
                if ($bc.CycleCount -gt 500) {
                    Write-StatusField "Cycles" "$($bc.CycleCount) cycles — high usage" -Status Warn
                    $warnings.Add("Battery has $($bc.CycleCount) charge cycles — high usage.")
                } elseif ($bc.CycleCount -gt 300) {
                    Write-StatusField "Cycles" "$($bc.CycleCount) cycles — moderate usage" -Status Warn
                } else {
                    Write-StatusField "Cycles" "$($bc.CycleCount) cycles — low usage" -Status OK
                }
            }
        }
    }

    # powercfg battery report
    if (-not $SkipBatteryReport) {
        $battReportPath = "$OutputPath-battery.html"
        try {
            $null = & powercfg /batteryreport /output $battReportPath 2>&1
            Write-Host "`n  📄 " -NoNewline
            Write-Host "Battery report: " -NoNewline -ForegroundColor Gray
            Write-Host $battReportPath -ForegroundColor Blue
        } catch {
            Write-Host "    ℹ  Could not generate battery report" -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host "  No battery detected (desktop or battery removed)" -ForegroundColor DarkGray
}

# ═══════════════════════════════════════════════════════════════════════════
#  6. DISPLAY / MONITOR
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "6. DISPLAY"

$monitors = Get-CimInstance WmiMonitorID -Namespace root\WMI -ErrorAction SilentlyContinue

if ($monitors) {
    $report['Display'] = [System.Collections.Generic.List[object]]::new()
    $monIndex = 0
    foreach ($mon in $monitors) {
        $monIndex++
        $mfr     = Decode-WmiString $mon.ManufacturerName
        $name    = Decode-WmiString $mon.UserFriendlyName
        $serial  = Decode-WmiString $mon.SerialNumberID
        $prodCode = Decode-WmiString $mon.ProductCodeID

        Write-Host ""
        Write-Field "  Monitor $monIndex" $(if ($name) { $name } else { "(internal)" })
        Write-Field "  Manufacturer Code" $mfr
        Write-Field "  Product Code" $prodCode
        Write-Field "  Serial" $serial
        Write-Field "  Year of Manufacture" $mon.YearOfManufacture
        Write-Field "  Week of Manufacture" $mon.WeekOfManufacture

        $isKnownPanel = Test-VendorMatch -ActualValue $mfr -KnownVendors $profile.PanelVendors
        if ($isKnownPanel) {
            Write-StatusField "  Panel Vendor" "'$mfr' is a known $($profile.DisplayName) panel supplier" -Status OK
        } elseif ($mfr) {
            Write-StatusField "  Panel Vendor" "'$mfr' — verify for your $($profile.DisplayName) model" -Status Warn
        }

        $report['Display'].Add([ordered]@{
            Name         = $name
            Manufacturer = $mfr
            ProductCode  = $prodCode
            Serial       = $serial
            YearOfMfr    = $mon.YearOfManufacture
            WeekOfMfr    = $mon.WeekOfManufacture
        })
    }
} else {
    Write-Host "  Could not retrieve monitor details" -ForegroundColor DarkGray
}

# GPU info
$videoControllers = Get-CimInstance Win32_VideoController
foreach ($vc in $videoControllers) {
    Write-Host ""
    Write-Field "  GPU" $vc.Name
    Write-Field "  Driver Version" $vc.DriverVersion
    Write-Field "  Resolution" "$($vc.CurrentHorizontalResolution) x $($vc.CurrentVerticalResolution)"
    if ($vc.AdapterRAM -and $vc.AdapterRAM -gt 0) {
        Write-Field "  Video RAM" (Get-FriendlySize $vc.AdapterRAM)
    }
}

# ═══════════════════════════════════════════════════════════════════════════
#  7. NETWORK ADAPTERS
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "7. NETWORK ADAPTERS"

$netAdapters = Get-CimInstance Win32_NetworkAdapter | Where-Object {
    $_.PhysicalAdapter -eq $true -and $_.PNPDeviceID -notlike "*ROOT*"
}

$report['Network'] = [System.Collections.Generic.List[object]]::new()

foreach ($nic in $netAdapters) {
    Write-Host ""
    Write-Field "  Adapter" $nic.Name
    Write-Field "  MAC Address" $nic.MACAddress
    Write-Field "  PNP Device ID" $nic.PNPDeviceID

    $isWifi = $nic.Name -match "Wi-Fi|Wireless|WLAN|WiFi|802\.11"
    if ($isWifi) {
        $isKnownVendor = Test-VendorMatch -ActualValue $nic.Name -KnownVendors $profile.WifiVendors
        if ($isKnownVendor) {
            Write-StatusField "  WiFi Vendor" "Known $($profile.DisplayName)-compatible wireless vendor" -Status OK
        } else {
            $warnMsg = "WiFi card vendor may not be $($profile.DisplayName)-standard"
            if ($profile.WifiWhitelist) {
                $warnMsg += " (this brand enforces BIOS WiFi whitelist!)"
            }
            Write-StatusField "  WiFi Vendor" $warnMsg -Status Warn
            $warnings.Add("WiFi adapter '$($nic.Name)' — verify with $($profile.DisplayName) approved list.")
        }
    }

    $report['Network'].Add([ordered]@{
        Name = $nic.Name
        MAC  = $nic.MACAddress
        PNPID = $nic.PNPDeviceID
        IsWifi = $isWifi
    })
}

# ═══════════════════════════════════════════════════════════════════════════
#  8. AUDIO
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "8. AUDIO"

$audio = Get-CimInstance Win32_SoundDevice
$report['Audio'] = [System.Collections.Generic.List[object]]::new()

foreach ($a in $audio) {
    Write-Field "Device" $a.Name
    Write-Field "Manufacturer" $a.Manufacturer
    $report['Audio'].Add([ordered]@{ Name = $a.Name; Manufacturer = $a.Manufacturer })
}

# ═══════════════════════════════════════════════════════════════════════════
#  9. CAMERA / WEBCAM
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "9. CAMERA / WEBCAM"

$cameras = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.PNPClass -eq 'Camera' -or $_.PNPClass -eq 'Image' -or $_.Name -match 'webcam|camera|IR'
}

$report['Camera'] = [System.Collections.Generic.List[object]]::new()

if ($cameras) {
    foreach ($cam in $cameras) {
        Write-Field "Camera" $cam.Name
        Write-Field "Device ID" $cam.DeviceID
        $report['Camera'].Add([ordered]@{ Name = $cam.Name; DeviceID = $cam.DeviceID })
    }
} else {
    Write-Host "  No camera detected" -ForegroundColor DarkGray
}

# ═══════════════════════════════════════════════════════════════════════════
#  10. AC ADAPTER / CHARGER
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "10. AC ADAPTER / CHARGER"

$acStatus = Get-CimInstance -Namespace root\WMI -ClassName BatteryStatus -ErrorAction SilentlyContinue
if ($acStatus) {
    foreach ($ac in $acStatus) {
        Write-Field "AC Connected" $(if ($ac.PowerOnline) { "Yes" } else { "No" })
        Write-Field "Charging" $(if ($ac.Charging) { "Yes" } else { "No" })
        if ($ac.Voltage) { Write-Field "Voltage" "$($ac.Voltage) mV" }
    }
}

# Check event log for adapter warnings
try {
    $adapterEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'System'; Level = 3
    } -MaxEvents 300 -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -like "*adapter*" -or $_.Message -like "*charger*" -or $_.Message -like "*power supply*" } |
    Select-Object -First 5

    if ($adapterEvents) {
        Write-StatusField "Adapter Alerts" "Found recent warnings in Event Log:" -Status Warn
        foreach ($evt in $adapterEvents) {
            $msgPreview = $evt.Message.Substring(0, [Math]::Min(100, $evt.Message.Length))
            Write-Host "     [$($evt.TimeCreated.ToString('yyyy-MM-dd'))] $msgPreview..." -ForegroundColor Yellow
        }
        $warnings.Add("AC adapter/charger warnings found in Windows Event Log.")
    } else {
        Write-StatusField "Adapter Status" "No charger warnings in Event Log" -Status OK
    }
} catch {
    Write-Host "  ℹ  Could not check Event Log" -ForegroundColor DarkGray
}

# ═══════════════════════════════════════════════════════════════════════════
#  11. TPM (TRUSTED PLATFORM MODULE)
# ═══════════════════════════════════════════════════════════════════════════

Write-Section "11. TPM (TRUSTED PLATFORM MODULE)"

try {
    $tpm = Get-CimInstance -Namespace root\CIMv2\Security\MicrosoftTpm -ClassName Win32_Tpm -ErrorAction SilentlyContinue
    if ($tpm) {
        Write-Field "TPM Present" "Yes"
        Write-Field "Manufacturer ID" $tpm.ManufacturerId
        Write-Field "Manufacturer" $tpm.ManufacturerIdTxt
        Write-Field "Spec Version" $tpm.SpecVersion
        Write-StatusField "TPM" "Module detected" -Status OK

        $report['TPM'] = [ordered]@{
            Present      = $true
            Manufacturer = $tpm.ManufacturerIdTxt
            SpecVersion  = $tpm.SpecVersion
        }
    } else {
        Write-Host "  Not detected or not accessible (requires admin)" -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  ℹ  Could not query TPM (requires admin)" -ForegroundColor DarkGray
}

# ═══════════════════════════════════════════════════════════════════════════
#  12. BRAND-SPECIFIC NOTES
# ═══════════════════════════════════════════════════════════════════════════

if ($profile.Notes -and $profile.Notes.Count -gt 0) {
    Write-Section "12. $($profile.DisplayName.ToUpper())-SPECIFIC NOTES"
    $noteNum = 0
    foreach ($note in $profile.Notes) {
        $noteNum++
        Write-Host "  $noteNum. $note" -ForegroundColor DarkYellow
    }
}

# ═══════════════════════════════════════════════════════════════════════════
#  FINAL SUMMARY
# ═══════════════════════════════════════════════════════════════════════════

Write-Banner "SUMMARY — $($profile.DisplayName.ToUpper())"

Write-Host ""
Write-Host "  System   : $detectedManufacturer $model" -ForegroundColor White
Write-Host "  Serial   : $serialNumber" -ForegroundColor White
Write-Host "  CPU      : $($cpus[0].Name.Trim())" -ForegroundColor White
Write-Host "  RAM      : $(Get-FriendlySize $totalRAM) ($($ramModules.Count) module(s))" -ForegroundColor White
Write-Host "  Storage  : $($disks.Count) drive(s)" -ForegroundColor White
Write-Host ""

if ($warnings.Count -eq 0) {
    Write-Host "  ✅ ALL CLEAR — No issues detected!" -ForegroundColor Green
    Write-Host "     All identified components appear to be from known $($profile.DisplayName) OEM suppliers." -ForegroundColor Green
    Write-Host "     Still recommended: compare with original config using serial '$serialNumber'" -ForegroundColor Gray
} else {
    Write-Host "  ⚠  $($warnings.Count) ITEM(S) FLAGGED FOR REVIEW:" -ForegroundColor Yellow
    Write-Host ""
    $warnNum = 0
    foreach ($w in $warnings) {
        $warnNum++
        Write-Host "     $warnNum. $w" -ForegroundColor Yellow
    }
}

if ($supportUrl) {
    Write-Host "`n  🔗 Verify original configuration:" -ForegroundColor Gray
    Write-Host "     $supportUrl" -ForegroundColor Blue
}

# ═══════════════════════════════════════════════════════════════════════════
#  EXPORT: TEXT REPORT
# ═══════════════════════════════════════════════════════════════════════════

$txtPath = "$OutputPath.txt"

$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine("=" * 64)
$null = $sb.AppendLine("  $($profile.DisplayName.ToUpper()) LAPTOP COMPONENT ORIGINALITY REPORT")
$null = $sb.AppendLine("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$null = $sb.AppendLine("  Brand     : $($profile.DisplayName)")
$null = $sb.AppendLine("  Serial    : $serialNumber")
$null = $sb.AppendLine("=" * 64)

foreach ($section in $report.Keys) {
    $null = $sb.AppendLine("`n--- $section ---")
    $data = $report[$section]

    if ($data -is [System.Collections.IList]) {
        $itemNum = 0
        foreach ($item in $data) {
            $itemNum++
            $null = $sb.AppendLine("  [$itemNum]")
            foreach ($key in $item.Keys) {
                $null = $sb.AppendLine("    {0,-26}: {1}" -f $key, $item[$key])
            }
        }
    } elseif ($data -is [hashtable] -or $data -is [System.Collections.Specialized.OrderedDictionary]) {
        foreach ($key in $data.Keys) {
            $null = $sb.AppendLine("  {0,-26}: {1}" -f $key, $data[$key])
        }
    }
}

if ($warnings.Count -gt 0) {
    $null = $sb.AppendLine("`n--- WARNINGS ($($warnings.Count)) ---")
    $wn = 0
    foreach ($w in $warnings) { $wn++; $null = $sb.AppendLine("  $wn. $w") }
} else {
    $null = $sb.AppendLine("`n--- NO ISSUES DETECTED ---")
}

if ($supportUrl) { $null = $sb.AppendLine("`nVerify original config: $supportUrl") }

$sb.ToString() | Out-File -FilePath $txtPath -Encoding utf8
Write-Host "`n  📄 Text report : $txtPath" -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════════════
#  EXPORT: JSON (optional)
# ═══════════════════════════════════════════════════════════════════════════

if ($ExportJson) {
    $jsonPath = "$OutputPath.json"
    $jsonExport = [ordered]@{
        GeneratedAt = (Get-Date -Format 'o')
        Brand       = $profile.DisplayName
        Serial      = $serialNumber
        Model       = $model
        Components  = $report
        Warnings    = $warnings
        SupportUrl  = $supportUrl
    }
    $jsonExport | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding utf8
    Write-Host "  📄 JSON report : $jsonPath" -ForegroundColor Green
}

# ═══════════════════════════════════════════════════════════════════════════
#  EXPORT: HTML (optional)
# ═══════════════════════════════════════════════════════════════════════════

if ($ExportHtml) {
    $htmlPath = "$OutputPath.html"

    $brandColor = switch ($Brand) {
        "Dell"   { "#007DB8" }
        "HP"     { "#0096D6" }
        "Lenovo" { "#E2231A" }
        default  { "#00d4ff" }
    }

    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>$($profile.DisplayName) Component Report — $serialNumber</title>
<style>
    body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #1a1a2e; color: #e0e0e0; margin: 0; padding: 2rem; }
    .header { background: linear-gradient(135deg, #16213e 0%, #0f3460 100%); padding: 2rem; border-radius: 12px; margin-bottom: 2rem; border-left: 6px solid $brandColor; }
    .header h1 { color: $brandColor; margin: 0 0 0.5rem 0; }
    .header p { margin: 0.2rem 0; color: #aaa; }
    h2 { color: $brandColor; margin-top: 2rem; border-bottom: 1px solid #333; padding-bottom: 0.3rem; }
    table { border-collapse: collapse; width: 100%; margin: 1rem 0; }
    th, td { padding: 8px 14px; text-align: left; border: 1px solid #333; }
    th { background: #16213e; color: $brandColor; font-weight: 600; }
    tr:nth-child(even) { background: #1a1a3e; }
    tr:hover { background: #222244; }
    .warn-box { background: #3e2723; border-left: 4px solid #ff9800; padding: 1rem 1.5rem; border-radius: 8px; margin: 1rem 0; }
    .ok-box { background: #1b3a1b; border-left: 4px solid #4caf50; padding: 1rem 1.5rem; border-radius: 8px; margin: 1rem 0; }
    .warn { color: #ffc107; } .ok { color: #4caf50; }
    a { color: #64b5f6; text-decoration: none; } a:hover { text-decoration: underline; }
    .meta { display: flex; gap: 2rem; flex-wrap: wrap; }
    .meta-item { background: #16213e; padding: 0.8rem 1.2rem; border-radius: 8px; min-width: 150px; }
    .meta-item .label { color: #888; font-size: 0.8rem; text-transform: uppercase; }
    .meta-item .value { color: #fff; font-size: 1.1rem; font-weight: 600; margin-top: 0.2rem; }
</style>
</head>
<body>
<div class="header">
    <h1>$($profile.DisplayName) Component Originality Report</h1>
    <div class="meta">
        <div class="meta-item"><div class="label">$($profile.SerialLabel)</div><div class="value">$serialNumber</div></div>
        <div class="meta-item"><div class="label">Model</div><div class="value">$model</div></div>
        <div class="meta-item"><div class="label">Generated</div><div class="value">$(Get-Date -Format 'yyyy-MM-dd HH:mm')</div></div>
    </div>
    $(if ($supportUrl) { "<p style='margin-top:1rem'><a href='$supportUrl' target='_blank'>🔗 View Original Configuration on $($profile.DisplayName) Support</a></p>" })
</div>
"@

    foreach ($section in $report.Keys) {
        $htmlContent += "<h2>$section</h2><table><tr><th>Property</th><th>Value</th></tr>"
        $data = $report[$section]

        if ($data -is [System.Collections.IList]) {
            foreach ($item in $data) {
                foreach ($key in $item.Keys) {
                    $htmlContent += "<tr><td>$key</td><td>$($item[$key])</td></tr>"
                }
                $htmlContent += "<tr><td colspan='2' style='border:none; height:8px;'></td></tr>"
            }
        } elseif ($data -is [hashtable] -or $data -is [System.Collections.Specialized.OrderedDictionary]) {
            foreach ($key in $data.Keys) {
                $htmlContent += "<tr><td>$key</td><td>$($data[$key])</td></tr>"
            }
        }
        $htmlContent += "</table>"
    }

    if ($warnings.Count -gt 0) {
        $htmlContent += "<div class='warn-box'><h2 class='warn'>⚠ Warnings ($($warnings.Count))</h2><ol>"
        foreach ($w in $warnings) { $htmlContent += "<li class='warn'>$w</li>" }
        $htmlContent += "</ol></div>"
    } else {
        $htmlContent += "<div class='ok-box'><h2 class='ok'>✅ All Clear</h2><p>All components appear to be from known $($profile.DisplayName) OEM suppliers.</p></div>"
    }

    $htmlContent += "<p style='color:#555; margin-top:2rem; font-size:0.85rem;'>Generated by Laptop Component Originality Checker</p></body></html>"
    $htmlContent | Out-File -FilePath $htmlPath -Encoding utf8
    Write-Host "  📄 HTML report : $htmlPath" -ForegroundColor Green
}

Write-Host "`n  Done! ✨`n" -ForegroundColor Cyan
