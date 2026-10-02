<#
    LaptopCheck - one report for laptop support

    Checks, and explains in plain English:
      1. Laptop identification
      2. Manufacturer support app - installed? If not, how to install it
      3. Windows version - still supported? How to update
      4. Drivers - updates waiting, devices with problems, where to download
      5. BIOS - installed version, newer one offered?, where to download
      6. Touchscreen - present when Windows was first set up? Working now?
         What to do (remove / update the touch driver)
      7. Screen - current panel, plus earlier panels this laptop has seen

    READ-ONLY. Nothing is installed, removed or changed.
    No administrator rights needed.

    Usage:
      .\LaptopCheck.ps1
      .\LaptopCheck.ps1 -Label "RMA-4471"
      .\LaptopCheck.ps1 -SkipWindowsUpdate      (faster, skips the online check)

    Data tables (Windows support dates, brand apps and links) reviewed 2026-10-01.
    Review the Windows table once a year, after each autumn Windows release.

    MIT Licensed. Provided as is, without warranty of any kind.
#>

[CmdletBinding()]
param(
    [string]$Label,
    [string]$OutputPath,
    [switch]$SkipWindowsUpdate,
    [switch]$NoPause
)

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

$ToolVersion  = '3.1'
$DataReviewed = '2026-10-01'

# =====================================================================
#  WINDOWS SUPPORT DATES
#  Source: Microsoft Lifecycle (Windows 11 Home and Pro / Enterprise and
#  Education), reviewed 2026-10-01. Last day of security updates.
# =====================================================================
$Win11Versions = @(
    [pscustomobject]@{ Build = 22000; Version = '21H2'; HomePro = '2023-10-10'; Enterprise = '2024-10-08' }
    [pscustomobject]@{ Build = 22621; Version = '22H2'; HomePro = '2024-10-08'; Enterprise = '2025-10-14' }
    [pscustomobject]@{ Build = 22631; Version = '23H2'; HomePro = '2025-11-11'; Enterprise = '2026-11-10' }
    [pscustomobject]@{ Build = 26100; Version = '24H2'; HomePro = '2026-10-13'; Enterprise = '2027-10-12' }
    [pscustomobject]@{ Build = 26200; Version = '25H2'; HomePro = '2027-10-12'; Enterprise = '2028-10-10' }
    [pscustomobject]@{ Build = 26300; Version = '26H2'; HomePro = '2028-10-09'; Enterprise = '2029-10-08' }
    [pscustomobject]@{ Build = 28000; Version = '26H1'; HomePro = '2028-03-14'; Enterprise = '2029-03-13' }
)
$Win10EndOfSupport = '2025-10-14'
$Win10EsuEnd       = '2027-10-12'   # consumer ESU, extended by Microsoft in June 2026

$Links = @{
    Win11Assistant = 'https://www.microsoft.com/software-download/windows11'
    PcHealthCheck  = 'https://aka.ms/GetPCHealthCheckApp'
    IntelDsa       = 'https://www.intel.com/content/www/us/en/support/detect.html'
    AmdDrivers     = 'https://www.amd.com/en/support/download/drivers.html'
}

# =====================================================================
#  BRANDS
#  Match      : regex against the manufacturer name reported by the BIOS
#  ModelMatch : optional extra regex against the model
#  Apps       : support apps; Detect is a regex checked against Start menu
#               names, installed-program names and Store package names
# =====================================================================
$BrandProfiles = @(
    @{
        Name = 'Lenovo'; Match = '^LENOVO'
        Apps = @(
            @{ Name = 'Lenovo Vantage';            Detect = '^Lenovo Vantage$|^E046963F\.LenovoCompanion$' }
            @{ Name = 'Lenovo Commercial Vantage'; Detect = '^Lenovo Commercial Vantage$|^E046963F\.LenovoSettingsforEnterprise$' }
            @{ Name = 'Lenovo System Update';      Detect = '^Lenovo System Update' }
        )
        Recommend = 'Lenovo Vantage'
        Install   = @('Microsoft Store: https://apps.microsoft.com/detail/9wzdncrfj4mv')
        HowToUse  = 'Open Lenovo Vantage and run its system update check.'
        Drivers   = 'https://pcsupport.lenovo.com'
        IdHint    = 'serial number'
    }
    @{
        Name = 'HP'; Match = '^HP\b|^Hewlett'
        Apps = @(
            @{ Name = 'HP Support Assistant'; Detect = '^HP Support Assistant|HPSupportAssistant' }
            @{ Name = 'HP Image Assistant';   Detect = '^HP Image Assistant' }
        )
        Recommend = 'HP Support Assistant'
        Install   = @('https://support.hp.com/us-en/help/hp-support-assistant')
        HowToUse  = 'Open HP Support Assistant and check for updates.'
        Drivers   = 'https://support.hp.com/us-en/drivers'
        IdHint    = 'serial number'
    }
    @{
        Name = 'Dell'; Match = '^Dell|^Alienware'
        Apps = @(
            @{ Name = 'Dell SupportAssist';    Detect = '^SupportAssist$|^Dell SupportAssist$|^DellInc\.DellSupportAssist' }
            @{ Name = 'Dell Command | Update'; Detect = '^Dell Command \| Update|^DellInc\.DellCommandUpdate' }
            @{ Name = 'Dell Update';           Detect = '^Dell Update' }
        )
        Recommend = 'Dell SupportAssist'
        Install   = @(
            'https://www.dell.com/en-us/lp/supportassist-for-home-pcs',
            'Business models (Latitude, Precision, Dell Pro) can use Dell Command | Update instead:',
            'https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update'
        )
        HowToUse  = 'Open SupportAssist (or Dell Command | Update) and check for updates.'
        Drivers   = 'https://www.dell.com/support/home/en-us?app=drivers'
        IdHint    = 'Service Tag'
    }
    @{
        Name = 'Acer'; Match = '^Acer|^Packard Bell|^Gateway'
        Apps = @(
            @{ Name = 'AcerSense';        Detect = 'AcerSense' }
            @{ Name = 'Acer Care Center'; Detect = 'Care ?Center' }
        )
        Recommend = 'AcerSense'
        Install   = @(
            'https://www.acer.com/us-en/support/drivers-and-manuals',
            'Enter your serial number or SNID, open "Application" and download AcerSense.'
        )
        HowToUse  = 'Open AcerSense (or Acer Care Center) and check for updates.'
        Drivers   = 'https://www.acer.com/us-en/support/drivers-and-manuals'
        IdHint    = 'serial number or SNID'
    }
    @{
        Name = 'ASUS'; Match = '^ASUS'
        Apps = @(
            @{ Name = 'MyASUS';        Detect = '^MyASUS|ASUSPCAssistant' }
            @{ Name = 'Armoury Crate'; Detect = 'Armoury ?Crate' }
        )
        Recommend = 'MyASUS'
        Install   = @(
            'Microsoft Store: https://apps.microsoft.com/detail/9n7r5s6b0zzh',
            'ROG and TUF gaming models: Armoury Crate - https://rog.asus.com/us/content/armoury-crate/'
        )
        HowToUse  = 'Open MyASUS (or Armoury Crate) and run its update check.'
        Drivers   = 'https://www.asus.com/support/'
        IdHint    = 'model name or serial number'
    }
    @{
        Name = 'MSI'; Match = '^Micro-Star|^MSI\b'
        Apps = @(
            @{ Name = 'MSI Center';    Detect = '^MSI Center|MSICenter' }
            @{ Name = 'Dragon Center'; Detect = 'Dragon ?Center' }
        )
        Recommend = 'MSI Center'
        Install   = @('Search your model on https://www.msi.com and download MSI Center from its support page.')
        HowToUse  = 'Open MSI Center > Support > Live Update > Scan.'
        Drivers   = 'https://www.msi.com/support/technical_details/NB_Driver_Update'
        IdHint    = 'model name'
    }
    @{
        Name = 'Samsung'; Match = '^SAMSUNG'
        Apps = @(
            @{ Name = 'Samsung Update'; Detect = 'Samsung ?Update' }
        )
        Recommend = 'Samsung Update'
        Install   = @('Microsoft Store: https://apps.microsoft.com/detail/9nq3hdb99vbf')
        HowToUse  = 'Open Samsung Update and install the updates it offers.'
        Drivers   = 'https://www.samsung.com/us/support/answer/ANS10001472/'
        IdHint    = 'model code'
    }
    @{
        Name = 'Microsoft Surface'; Match = '^Microsoft'; ModelMatch = 'Surface'
        Apps = @()
        NoAppNote = 'Surface drivers and firmware are delivered through Windows Update.'
        Drivers   = 'https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120'
        IdHint    = 'Surface model'
    }
    @{
        Name = 'Razer'; Match = '^Razer'
        Apps = @()
        NoAppNote = 'Razer publishes laptop drivers and BIOS updaters on its support site.'
        Drivers   = 'https://mysupport.razer.com'
        IdHint    = 'model or serial number'
    }
    @{
        Name = 'GIGABYTE'; Match = '^GIGABYTE|^AORUS'
        Apps = @(
            @{ Name = 'GIGABYTE Control Center'; Detect = 'GIGABYTE Control Center' }
        )
        Recommend = 'GIGABYTE Control Center'
        Install   = @('https://www.gigabyte.com/us/Support/Consumer/Download')
        HowToUse  = 'Open GIGABYTE Control Center and check for updates.'
        Drivers   = 'https://www.gigabyte.com/us/Support/Consumer/Download'
        IdHint    = 'model name'
    }
    @{
        Name = 'LG'; Match = '^LG\b'
        Apps = @(
            @{ Name = 'LG Update'; Detect = '^LG Update' }
        )
        Recommend = 'LG Update'
        Install   = @('Microsoft Store (LG Update Installer): https://apps.microsoft.com/detail/9ns1bb1x0d4k')
        HowToUse  = 'Open LG Update and install the updates it offers.'
        Drivers   = 'https://www.lg.com/us/support/help-library/how-to-install-and-update-drivers-on-your-lg-laptop--20155410253523'
        IdHint    = 'model name'
    }
    @{
        Name = 'HUAWEI'; Match = '^HUAWEI'
        Apps = @(
            @{ Name = 'HUAWEI PC Manager'; Detect = '^(HUAWEI )?PC Manager$' }
        )
        Recommend = 'HUAWEI PC Manager'
        Install   = @('https://consumer.huawei.com/en/support/content/en-us00688514/')
        HowToUse  = 'Open PC Manager and install the driver updates it offers.'
        Drivers   = 'https://consumer.huawei.com/en/support/driver-list/'
        IdHint    = 'serial number'
    }
    @{
        Name = 'HONOR'; Match = '^HONOR'
        Apps = @(
            @{ Name = 'HONOR PC Manager'; Detect = '^(HONOR )?PC Manager$' }
        )
        Recommend = 'HONOR PC Manager'
        Install   = @('https://www.honor.com/global/support/content/en-us15815743/')
        HowToUse  = 'Open PC Manager and install the driver updates it offers.'
        Drivers   = 'https://www.honor.com/global/support/content/en-us15815743/'
        IdHint    = 'serial number'
    }
    @{
        Name = 'Xiaomi'; Match = '^Xiaomi|^TIMI'
        Apps = @()
        NoAppNote = 'Xiaomi publishes laptop drivers on its regional support site.'
        Drivers   = 'https://www.mi.com/in/service/support/laptop-drivers.html'
        IdHint    = 'model name'
    }
    @{
        Name = 'Dynabook / Toshiba'; Match = '^Dynabook|^TOSHIBA'
        Apps = @(
            @{ Name = 'dynabook Service Station'; Detect = 'Service Station' }
        )
        Recommend = 'dynabook Service Station'
        Install   = @('https://support.dynabook.com/support/viewContentDetail?contentId=4016624')
        HowToUse  = 'Open Service Station and install the updates it offers.'
        Drivers   = 'https://support.dynabook.com'
        IdHint    = 'serial number'
    }
    @{
        Name = 'Fujitsu'; Match = '^FUJITSU'
        Apps = @(
            @{ Name = 'Fujitsu DeskUpdate'; Detect = 'DeskUpdate' }
        )
        Recommend = 'Fujitsu DeskUpdate'
        Install   = @('https://download.ts.fujitsu.com/deskupdate/index.asp?lng=EN')
        HowToUse  = 'Run DeskUpdate and install the driver updates it offers.'
        Drivers   = 'https://download.ts.fujitsu.com/deskupdate/index.asp?lng=EN'
        IdHint    = 'serial number'
    }
    @{
        Name = 'Panasonic'; Match = '^Panasonic'
        Apps = @()
        NoAppNote = 'Panasonic publishes TOUGHBOOK drivers and BIOS on its support site.'
        Drivers   = 'https://pc-dl.panasonic.co.jp/dl/search'
        IdHint    = 'model number'
    }
    @{
        Name = 'VAIO'; Match = '^VAIO|^Sony'
        Apps = @(
            @{ Name = 'VAIO Update'; Detect = 'VAIO (Update|Control Center)' }
        )
        Recommend = 'VAIO Update'
        Install   = @('https://support.us.vaio.com')
        HowToUse  = 'Open VAIO Update and install the updates it offers.'
        Drivers   = 'https://support.us.vaio.com'
        IdHint    = 'model number'
    }
    @{
        Name = 'Framework'; Match = '^Framework'
        Apps = @()
        NoAppNote = 'Framework provides a Windows driver bundle and BIOS updates on its knowledge base.'
        Drivers   = 'https://knowledgebase.frame.work/bios-and-drivers-downloads-rJ3PaCexh'
        IdHint    = 'laptop model and processor'
    }
    @{
        Name = 'Clevo / TongFang based'
        Match = '^(Notebook|TongFang|Schenker|Eluktronics|Sager|XMG|PCSpecialist|Metabox|Clevo)'
        Apps = @()
        NoAppNote = 'This laptop is built on a Clevo or TongFang chassis. Drivers come from the company that sold it (for example XMG, Eluktronics, Sager, PCSpecialist, Metabox).'
        Drivers   = ''
        IdHint    = 'chassis code'
    }
)

$GenericProfile = @{
    Name      = 'unknown manufacturer'
    Apps      = @()
    NoAppNote = 'No support app is known for this manufacturer.'
    Drivers   = ''
    IdHint    = 'serial number'
}

$ProblemCodes = @{
    1  = 'not configured correctly'
    3  = 'driver may be corrupted'
    10 = 'cannot start'
    12 = 'not enough resources'
    14 = 'needs a restart'
    18 = 'driver needs reinstalling'
    19 = 'registry entry damaged'
    21 = 'being removed by Windows'
    22 = 'DISABLED'
    24 = 'not present or missing driver'
    28 = 'NO DRIVER INSTALLED'
    31 = 'driver failed to load'
    32 = 'driver disabled'
    37 = 'driver failed to start'
    39 = 'driver corrupted or missing'
    43 = 'stopped by Windows (reported a fault)'
    45 = 'not connected'
    52 = 'driver signature problem'
    99 = 'problem (exact code unknown)'
}

# =====================================================================
#  SMALL HELPERS
# =====================================================================
function ConvertTo-Date {
    param([string]$Text)
    [datetime]::ParseExact($Text, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-Date {
    # InvariantCulture: some regional calendars (Thai, Arabic...) would print a different year
    param($Value)
    if ($null -eq $Value) { return 'unknown' }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) } catch { return 'unknown' }
}

function Get-ProblemText {
    param([int]$Code)
    if ($ProblemCodes.ContainsKey($Code)) { return "code $Code - $($ProblemCodes[$Code])" }
    return "code $Code"
}

function Get-ShortId {
    # 'PCI\VEN_8086&DEV_51E8&SUBSYS_...\3&11583659&0&A0' -> 'PCI\VEN_8086&DEV_51E8&SUBSYS_...'
    param([string]$InstanceId)
    $parts = @($InstanceId -split '\\')
    if ($parts.Count -ge 2) { return ('{0}\{1}' -f $parts[0], $parts[1]) }
    return $InstanceId
}

function Test-DevicePresent {
    # Get-PnpDevice reports Present on Windows 10/11. If it is ever missing,
    # fall back to Status (devices that are not plugged in report 'Unknown').
    param($Device)
    if ($null -ne $Device.Present) { return [bool]$Device.Present }
    return ("$($Device.Status)" -ne 'Unknown')
}

function Get-ErrorCode {
    # Device Manager error code as a number (0 = no problem). Works whether Windows
    # hands back a number, an enum, or a name such as CM_PROB_NONE / CM_PROB_FAILED_START.
    param($Device)
    $c = $Device.ConfigManagerErrorCode
    if ($null -eq $c) { return 0 }
    try { return [int]$c } catch { }
    if ("$c" -match '^\s*(CM_PROB_NONE)?\s*$') { return 0 }
    return 99        # some problem, exact code unknown
}

function Get-UpdateKind {
    # Sorts one Windows Update result into Ignore / Firmware / Driver / Feature / Software
    param([string]$Title, [int]$Type, [string]$DriverClass = '')
    if ($Type -eq 2) {
        if ($DriverClass -match 'Firmware' -or $Title -match 'Firmware|BIOS|UEFI') { return 'Firmware' }
        return 'Driver'
    }
    # Antivirus definitions and the malware removal tool are always "waiting"; not worth a customer action
    if ($Title -match 'Security Intelligence Update|Definition Update|Malicious Software Removal') { return 'Ignore' }
    if ($Title -match 'Feature update|Windows 11, version|Upgrade to Windows 11') { return 'Feature' }
    return 'Software'
}

# =====================================================================
#  EDID DECODER
#  Returns PnP ID, part number (if the panel stores one), native mode,
#  highest refresh rate and physical size.
# =====================================================================
function Get-EdidTiming {
    param([byte[]]$Bytes, [int]$Offset)
    $e = $Bytes; $o = $Offset
    $px = [int]$e[$o] -bor ([int]$e[$o + 1] -shl 8)
    # 0 = this slot is a text descriptor. A tiny value such as 1 is a placeholder that newer panels
    # use to say "the real timings are in the DisplayID block". Anything under 10 MHz (stored value
    # under 1000) is not a real laptop timing either: edid-decode also treats it as invalid data.
    if ($px -lt 1000) { return $null }
    $hAct = [int]$e[$o + 2] -bor ((([int]$e[$o + 4] -shr 4) -band 0x0F) -shl 8)
    $hBlk = [int]$e[$o + 3] -bor ((( [int]$e[$o + 4])        -band 0x0F) -shl 8)
    $vAct = [int]$e[$o + 5] -bor ((([int]$e[$o + 7] -shr 4) -band 0x0F) -shl 8)
    $vBlk = [int]$e[$o + 6] -bor ((( [int]$e[$o + 7])        -band 0x0F) -shl 8)
    $hTot = $hAct + $hBlk
    $vTot = $vAct + $vBlk
    if ($hTot -le 0 -or $vTot -le 0) { return $null }
    # Pixel clock is stored in units of 10 kHz
    $hz = [math]::Round(($px * 10000.0) / ($hTot * $vTot), 0)
    # Physical image size in millimetres (bytes 12-14 of the descriptor). Only meaningful in a
    # base-block timing; extension-block timings may leave it empty.
    $wMm = [int]$e[$o + 12] -bor ((([int]$e[$o + 14] -shr 4) -band 0x0F) -shl 8)
    $hMm = [int]$e[$o + 13] -bor ((( [int]$e[$o + 14])        -band 0x0F) -shl 8)
    [pscustomobject]@{ Width = $hAct; Height = $vAct; Refresh = $hz; WidthMm = $wMm; HeightMm = $hMm }
}

function Get-DisplayIdTimings {
    # DisplayID extension block (tag 0x70). Many newer laptop panels list their high-refresh modes
    # ONLY here: Type I detailed timings (DisplayID 1.x) or Type VII (DisplayID 2.0).
    # Without this a 165 Hz panel would be reported as 60 Hz.
    param([byte[]]$Bytes, [int]$Base)
    $e = $Bytes
    $list = New-Object System.Collections.Generic.List[object]
    $isV2 = ([int]$e[$Base + 1] -ge 0x20)            # 0x13 = DisplayID 1.3, 0x20 = DisplayID 2.0
    $end  = $Base + 5 + [int]$e[$Base + 2]           # first byte after the data blocks
    if ($end -gt $Base + 126) { $end = $Base + 126 }
    $p = $Base + 5
    while ($p + 3 -le $end) {
        $tag = [int]$e[$p]; $rev = [int]$e[$p + 1]; $len = [int]$e[$p + 2]
        if ($p + 3 + $len -gt $end) { break }        # damaged block
        $size = 0
        if (-not $isV2 -and $tag -eq 0x03) { $size = 20 }
        if ($isV2 -and $tag -eq 0x22)      { $size = 20 + (($rev -shr 4) -band 0x07) }
        if ($size -gt 0) {
            for ($q = $p + 3; $q + $size -le $p + 3 + $len; $q += $size) {
                $clk  = [int]$e[$q] -bor ([int]$e[$q + 1] -shl 8) -bor ([int]$e[$q + 2] -shl 16)
                $hAct = 1 + ([int]$e[$q + 4]  -bor ([int]$e[$q + 5]  -shl 8))
                $hBlk = 1 + ([int]$e[$q + 6]  -bor ([int]$e[$q + 7]  -shl 8))
                $vAct = 1 + ([int]$e[$q + 12] -bor ([int]$e[$q + 13] -shl 8))
                $vBlk = 1 + ([int]$e[$q + 14] -bor ([int]$e[$q + 15] -shl 8))
                # Type I counts the pixel clock in 10 kHz steps, Type VII in 1 kHz steps (both stored minus 1)
                $kHz  = if ($isV2) { $clk + 1 } else { 10 * ($clk + 1) }
                $hz   = [math]::Round(($kHz * 1000.0) / (($hAct + $hBlk) * ($vAct + $vBlk)), 0)
                $list.Add([pscustomobject]@{ Width = $hAct; Height = $vAct; Refresh = $hz; WidthMm = 0; HeightMm = 0 })
            }
        }
        $p += 3 + $len
    }
    $list.ToArray()
}

function ConvertFrom-EdidBytes {
    param([byte[]]$Bytes)
    $e = $Bytes
    if (-not $e -or $e.Length -lt 128) { return $null }
    if ($e[0] -ne 0 -or $e[1] -ne 255 -or $e[7] -ne 0) { return $null }

    # Manufacturer ID: three 5-bit letters, big endian ('A' = 1)
    $v = ([int]$e[8] -shl 8) -bor [int]$e[9]
    $letters = [char[]]@(
        ((($v -shr 10) -band 0x1F) + 64),
        ((($v -shr 5)  -band 0x1F) + 64),
        (( $v          -band 0x1F) + 64)
    )
    $vendor  = -join $letters
    $product = [int]$e[10] -bor ([int]$e[11] -shl 8)
    $pnpId   = '{0}{1:X4}' -f $vendor, $product

    $timings = New-Object System.Collections.Generic.List[object]
    $strings = New-Object System.Collections.Generic.List[string]

    foreach ($o in 54, 72, 90, 108) {
        $t = Get-EdidTiming -Bytes $e -Offset $o
        if ($t) { $timings.Add($t); continue }
        if ([int]$e[$o] -ne 0 -or [int]$e[$o + 1] -ne 0) { continue }    # starts with a pixel clock: placeholder or damaged slot, not text
        # 0xFE = text (where laptop panels keep the part number), 0xFC = name.
        # 0xFF is the serial number and must not be mistaken for a part number.
        $tag = [int]$e[$o + 3]
        if ($tag -eq 0xFC -or $tag -eq 0xFE) {
            $sb = New-Object System.Text.StringBuilder
            for ($i = 5; $i -le 17; $i++) {
                $c = [int]$e[$o + $i]
                if ($c -eq 0x0A -or $c -eq 0) { break }
                [void]$sb.Append([char]$c)
            }
            $s = $sb.ToString().Trim()
            if ($s) { $strings.Add($s) }
        }
    }

    # Extra timings in extension blocks: CTA-861 (tag 0x02) and DisplayID (tag 0x70) - high refresh modes often live here
    $extCount = [int]$e[126]
    for ($x = 1; $x -le $extCount; $x++) {
        $base = $x * 128
        if ($e.Length -lt $base + 128) { break }
        if ([int]$e[$base] -eq 0x70) {
            foreach ($t in @(Get-DisplayIdTimings -Bytes $e -Base $base)) { $timings.Add($t) }
            continue
        }
        if ([int]$e[$base] -ne 0x02) { continue }
        $dtd = [int]$e[$base + 2]
        if ($dtd -lt 4) { continue }
        $p = $base + $dtd
        while ($p + 18 -le $base + 127) {
            $t = Get-EdidTiming -Bytes $e -Offset $p
            if (-not $t) { break }
            $timings.Add($t)
            $p += 18
        }
    }

    # Part number, when the panel stores one in a text descriptor.
    # Real panel part numbers always contain a digit; plain words (SAMSUNG, DISPLAY) are not part numbers.
    $part = ''
    foreach ($s in $strings) {
        $u = $s.ToUpperInvariant()      # not ToUpper(): in a Turkish locale i becomes a different letter
        if ($u -match '^[A-Z0-9][A-Z0-9\.\-]{5,}$' -and
            $u -match '[0-9]' -and
            $u -notmatch '^(LCD|LED|MONITOR|DISPLAY|GENERIC|COLOR)') { $part = $s; break }
    }

    $native = $null
    if ($timings.Count -gt 0) { $native = $timings[0] }
    $maxHz = 0
    foreach ($t in $timings) { if ($t.Refresh -gt $maxHz) { $maxHz = $t.Refresh } }

    # Size: the first timing carries the image size in mm (exact). The header only has whole cm.
    $wCm = [int]$e[21]; $hCm = [int]$e[22]
    $wMm = 0; $hMm = 0
    if ($native -and $native.WidthMm -gt 0 -and $native.HeightMm -gt 0) { $wMm = $native.WidthMm; $hMm = $native.HeightMm }
    $diag = 0
    if ($wMm -gt 0) {
        $diag = [math]::Round([math]::Sqrt(($wMm * $wMm) + ($hMm * $hMm)) / 25.4, 1)
    }
    elseif ($wCm -gt 0 -and $hCm -gt 0) {
        $diag = [math]::Round([math]::Sqrt(($wCm * $wCm) + ($hCm * $hCm)) / 2.54, 1)
    }

    $aspect = ''
    if ($native -and $native.Height -gt 0) {
        $r = $native.Width / $native.Height
        if     ($r -gt 1.76 -and $r -lt 1.79) { $aspect = '16:9' }
        elseif ($r -gt 1.59 -and $r -lt 1.61) { $aspect = '16:10' }
        elseif ($r -gt 1.49 -and $r -lt 1.51) { $aspect = '3:2' }
    }

    [pscustomobject]@{
        PnpId    = $pnpId
        Part     = $part
        Strings  = @($strings)
        Width    = if ($native) { $native.Width }  else { 0 }
        Height   = if ($native) { $native.Height } else { 0 }
        Aspect   = $aspect
        MaxHz    = $maxHz
        WidthCm  = $wCm
        WidthMm  = $wMm
        HeightMm = $hMm
        Diagonal = $diag
        Built    = "week $([int]$e[16]), $(1990 + [int]$e[17])"
    }
}

# =====================================================================
#  DECISIONS (no system access - easy to test)
# =====================================================================
function Get-BrandProfile {
    param([string]$Manufacturer, [string]$Model, [string]$BoardMaker = '')

    # Cheap and white-box laptops often leave the BIOS maker name as a placeholder
    $placeholder = '^\s*$|^(To Be Filled|System manufacturer|Default string|Not Applicable|Not Specified|N/A|None|Unknown|O\.?E\.?M\.?$)'
    $name = $Manufacturer
    if ($name -match $placeholder) { $name = $BoardMaker }      # only then look at the motherboard maker
    if ($name -match $placeholder) { $name = '' }

    if ($name) {
        foreach ($b in $BrandProfiles) {
            if ($name -match $b.Match) {
                if ($b.ModelMatch -and $Model -notmatch $b.ModelMatch) { continue }
                return $b
            }
        }
    }
    $g = $GenericProfile.Clone()
    if ($name) { $g.Name = $name.Trim() }
    return $g
}

function Find-SupportApp {
    param($Brand, [string[]]$AppNames)
    $found = @()
    foreach ($app in @($Brand.Apps)) {
        $hit = @($AppNames | Where-Object { $_ -match $app.Detect })
        if ($hit.Count -gt 0) { $found += $app.Name }
    }
    return $found
}

function Get-WindowsVerdict {
    param([int]$Build, [string]$EditionId, [datetime]$Today)

    $r = [ordered]@{
        Product = ''; Version = ''; EndDate = $null; DaysLeft = $null
        State = 'Unknown'; Enterprise = $false
    }
    $r.Enterprise = ($EditionId -match '^(Enterprise|Education|IoTEnterprise)N?$|ServerRdsh')
    $isLtsc       = ($EditionId -match 'EnterpriseS|IoTEnterpriseS')

    if ($Build -ge 22000) {
        $r.Product = 'Windows 11'
        $row = $Win11Versions | Where-Object { $_.Build -eq $Build } | Select-Object -First 1
        if ($row) { $r.Version = $row.Version }
        if ($isLtsc) { $r.State = 'LTSC'; return [pscustomobject]$r }
        if (-not $row) {
            # Compare against the annual (H2) releases; 26H1 uses a separate, higher build line
            $newest = ($Win11Versions | Where-Object { $_.Version -like '*H2' } |
                       Measure-Object -Property Build -Maximum).Maximum
            if ($Build -gt $newest) { $r.State = 'Newer' }
            return [pscustomobject]$r
        }
        $endText = if ($r.Enterprise) { $row.Enterprise } else { $row.HomePro }
        $r.EndDate  = ConvertTo-Date $endText
        $r.DaysLeft = [int][math]::Floor(($r.EndDate - $Today.Date).TotalDays)
        if     ($r.DaysLeft -lt 0)  { $r.State = 'Ended' }
        elseif ($r.DaysLeft -le 60) { $r.State = 'EndingSoon' }
        else                        { $r.State = 'Supported' }
    }
    elseif ($Build -ge 10240) {
        $r.Product = 'Windows 10'
        $r.Version = if ($Build -eq 19045) { '22H2' } else { "build $Build" }
        $r.EndDate = ConvertTo-Date $Win10EndOfSupport
        $r.State   = if ($Build -eq 19045) { 'Win10' } else { 'Win10Old' }
        if ($isLtsc) { $r.State = 'LTSC' }
    }
    else {
        $r.Product = 'Windows (older than Windows 10)'
        $r.State   = 'Ended'
    }
    return [pscustomobject]$r
}

function Get-TouchVerdict {
    param($TouchDevices, $OriginalInstall, [bool]$DriversMissing)

    # @($null) is a one-item array in PowerShell, so drop empty entries first
    $devs = @($TouchDevices | Where-Object { $null -ne $_ })
    $r = [ordered]@{ History = 'Never'; FirstSeen = $null; Now = 'None'; LastSeen = $null; Code = 0 }

    if ($devs.Count -eq 0) {
        if ($DriversMissing) { $r.History = 'CannotTell' }
        return [pscustomobject]$r
    }

    # When did a touchscreen first appear?
    $first = $null
    foreach ($d in $devs) {
        $dt = $d.FirstInstall
        if (-not $dt) { $dt = $d.InstallDate }
        if ($dt -and (-not $first -or $dt -lt $first)) { $first = $dt }
    }
    $r.FirstSeen = $first
    if (-not $first -or -not $OriginalInstall) {
        # A touch device exists, but Windows gave no date to compare: do not claim it came with the laptop
        $r.History = 'NoDate'
    } elseif ($first -le ([datetime]$OriginalInstall).AddDays(3)) {
        $r.History = 'Yes'
    } else {
        $r.History = 'AddedLater'
    }

    # What is it doing now?
    $present = @($devs | Where-Object { $_.Present })
    if ($present.Count -gt 0) {
        $bad = @($present | Where-Object { [int]$_.ErrorCode -ne 0 }) | Select-Object -First 1
        if (-not $bad) {
            $r.Now = 'Working'
        } elseif ([int]$bad.ErrorCode -eq 22) {
            $r.Now = 'Disabled'; $r.Code = 22
        } else {
            $r.Now = 'Error'; $r.Code = [int]$bad.ErrorCode
        }
    } else {
        $r.Now = 'NotConnected'
        foreach ($d in $devs) {
            $seen = $d.LastRemoval
            if (-not $seen) { $seen = $d.LastArrival }
            if ($seen -and (-not $r.LastSeen -or $seen -gt $r.LastSeen)) { $r.LastSeen = $seen }
        }
    }
    return [pscustomobject]$r
}

function Select-Screens {
    # Splits recorded screens into the one in use now and earlier ones.
    param($Screens)
    $all = @($Screens | Where-Object { $_.Edid })
    $laptopSize = { param($s) $s.Edid.WidthCm -ge 20 -and $s.Edid.WidthCm -le 42 }

    $now = @($all | Where-Object { $_.Present -and (& $laptopSize $_) } |
             Sort-Object { $_.Edid.WidthCm }) | Select-Object -First 1
    if (-not $now) {
        # Some built-in panels report no physical size; external monitors always do
        $now = @($all | Where-Object { $_.Present -and $_.Edid.WidthCm -eq 0 }) | Select-Object -First 1
    }

    $earlier = @()
    $seen = @{}
    if ($now) { $seen[$now.Edid.PnpId] = $true }
    foreach ($s in @($all | Where-Object { -not $_.Present -and (& $laptopSize $_) } |
                     Sort-Object { if ($_.LastArrival) { [datetime]$_.LastArrival } else { [datetime]::MinValue } } -Descending)) {
        if ($seen.ContainsKey($s.Edid.PnpId)) { continue }
        $seen[$s.Edid.PnpId] = $true
        $earlier += $s
    }
    [pscustomobject]@{ Now = $now; Earlier = $earlier }
}

# =====================================================================
#  COLLECTORS (read from Windows)
# =====================================================================
function Write-Step {
    param([string]$Text)
    Write-Host "  $Text" -ForegroundColor DarkGray
}

function Get-DeviceDate {
    param([string]$InstanceId, [string]$Key)
    try {
        $d = (Get-PnpDeviceProperty -InstanceId $InstanceId -KeyName $Key -ErrorAction Stop).Data
        if ($d -is [datetime] -and $d.Year -ge 2000) { return $d }
    } catch { }
    return $null
}

function Get-FriendlyModel {
    # Lenovo keeps the friendly model name in Version (Model is the machine type, e.g. 83DL)
    param($ComputerSystem, $Product)
    $friendly = $ComputerSystem.Model
    if ($ComputerSystem.Manufacturer -match '^LENOVO' -and $Product.Version -and
        $Product.Version -notmatch '^(None|Not|System|To Be|Default)') { $friendly = $Product.Version }
    return $friendly
}

function Find-TouchDevices {
    # Touchscreen = HID collection with usage page 0x0D (digitizer), usage 0x04 (touch screen).
    # This works on any Windows language; the English name is only a fallback.
    # Includes devices that are remembered but not connected: check Present on each.
    param($Devices)
    $touchId = 'HID_DEVICE_UP:000D_U:0004'
    $found = @($Devices | Where-Object {
        $_.Class -eq 'HIDClass' -and (
            (@($_.CompatibleID) -contains $touchId) -or ("$($_.FriendlyName)" -match 'touch ?screen')
        )
    })
    foreach ($d in @($Devices | Where-Object { $_.Class -eq 'HIDClass' -and -not $_.CompatibleID })) {
        try {
            $ids = (Get-PnpDeviceProperty -InstanceId $d.InstanceId -KeyName 'DEVPKEY_Device_CompatibleIds' -ErrorAction Stop).Data
            if (@($ids) -contains $touchId -and @($found | Where-Object { $_.InstanceId -eq $d.InstanceId }).Count -eq 0) {
                $found += $d
            }
        } catch { }
    }
    return @($found)
}

function Get-ScreenList {
    # Every screen Windows has a stored EDID for: the one connected now and ones seen before.
    param($Devices)
    $present = @{}
    foreach ($m in @($Devices | Where-Object { $_.Class -eq 'Monitor' })) {
        $present[([string]$m.InstanceId).ToUpper()] = (Test-DevicePresent $m)
    }
    $screens = @()
    foreach ($pnp in @(Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Enum\DISPLAY' -ErrorAction SilentlyContinue)) {
        foreach ($inst in @(Get-ChildItem -LiteralPath $pnp.PSPath -ErrorAction SilentlyContinue)) {
            $raw = (Get-ItemProperty -LiteralPath (Join-Path $inst.PSPath 'Device Parameters') -ErrorAction SilentlyContinue).EDID
            if (-not $raw) { continue }
            $info = ConvertFrom-EdidBytes -Bytes ([byte[]]$raw)
            if (-not $info) { continue }
            $id = ('DISPLAY\{0}\{1}' -f $pnp.PSChildName, $inst.PSChildName).ToUpper()
            $isPresent = $false
            if ($present.ContainsKey($id)) { $isPresent = $present[$id] }
            $screens += [pscustomobject]@{
                InstanceId  = $id
                Present     = $isPresent
                LastArrival = Get-DeviceDate $id 'DEVPKEY_Device_LastArrivalDate'
                Edid        = $info
            }
        }
    }
    return @($screens)
}

function Get-OriginalInstallDate {
    # Win32_OperatingSystem.InstallDate moves forward with every feature
    # update. Earlier dates survive under HKLM\SYSTEM\Setup\Source OS*.
    $dates = @()
    $os = Get-CimInstance Win32_OperatingSystem
    if ($os.InstallDate) { $dates += [datetime]$os.InstallDate }
    foreach ($k in @(Get-ChildItem 'HKLM:\SYSTEM\Setup' -ErrorAction SilentlyContinue |
                     Where-Object { $_.PSChildName -like 'Source OS*' })) {
        $epoch = (Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue).InstallDate
        if ($epoch) {
            try { $dates += [DateTimeOffset]::FromUnixTimeSeconds([int64]$epoch).LocalDateTime } catch { }
        }
    }
    $dates | Sort-Object | Select-Object -First 1
}

function Get-InstalledAppNames {
    $names = New-Object System.Collections.Generic.List[string]
    try {
        # No -ErrorAction here: in Windows PowerShell 5.1 Get-StartApps is a plain function that only has -Name,
        # so -ErrorAction would be rejected and the Start menu names would silently never be read.
        foreach ($a in @(Get-StartApps)) { if ($a.Name) { $names.Add([string]$a.Name) } }
    } catch { }
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($k in $keys) {
        foreach ($i in @(Get-ItemProperty -Path $k -ErrorAction SilentlyContinue)) {
            if ($i.DisplayName) { $names.Add([string]$i.DisplayName) }
        }
    }
    try {
        foreach ($p in @(Get-AppxPackage -ErrorAction Stop)) { if ($p.Name) { $names.Add([string]$p.Name) } }
    } catch { }
    $names.ToArray()
}

function Get-WindowsUpdateFacts {
    $f = [ordered]@{
        Searched = $false; Error = ''
        Software = @(); Feature = @(); Drivers = @(); Firmware = @()
    }
    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $result   = $searcher.Search('IsInstalled=0 and IsHidden=0')
        $f.Searched = $true
        foreach ($u in $result.Updates) {
            $title = [string]$u.Title
            $class = ''
            if ([int]$u.Type -eq 2) { try { $class = [string]$u.DriverClass } catch { } }
            $item = [pscustomobject]@{ Title = $title; Class = $class; Optional = [bool]$u.BrowseOnly }
            switch (Get-UpdateKind -Title $title -Type ([int]$u.Type) -DriverClass $class) {
                'Firmware' { $f.Firmware += $item }
                'Driver'   { $f.Drivers  += $item }
                'Feature'  { $f.Feature  += $title }
                'Software' { $f.Software += $title }
            }
        }
    } catch {
        # A failed COM call arrives wrapped as: Exception calling "Search" with "1" argument(s): "..."
        # The inner message ("Exception from HRESULT: 0x8024402C") is the useful part.
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException -and $_.Exception.InnerException.Message) { $msg = $_.Exception.InnerException.Message }
        $f.Error = $msg
    }
    [pscustomobject]$f
}

function Get-AllFacts {
    Write-Step 'Reading laptop details...'
    $cs   = Get-CimInstance Win32_ComputerSystem
    $csp  = Get-CimInstance Win32_ComputerSystemProduct
    $bios = Get-CimInstance Win32_BIOS
    $bb   = Get-CimInstance Win32_BaseBoard
    $os   = Get-CimInstance Win32_OperatingSystem
    $cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
    $cv   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

    $friendly = Get-FriendlyModel -ComputerSystem $cs -Product $csp

    $lastFix = Get-HotFix -ErrorAction SilentlyContinue | Where-Object { $_.InstalledOn } |
               Sort-Object InstalledOn -Descending | Select-Object -First 1

    Write-Step 'Checking installed support apps...'
    $apps = @(Get-InstalledAppNames)

    $wu = $null
    if (-not $SkipWindowsUpdate) {
        Write-Step 'Asking Windows Update what is available (can take a few minutes)...'
        $wu = Get-WindowsUpdateFacts
    }

    Write-Step 'Checking devices and drivers...'
    $all = @(Get-PnpDevice -ErrorAction SilentlyContinue)

    $problems = @($all | Where-Object {
        (Test-DevicePresent $_) -and ((Get-ErrorCode $_) -ne 0)
    } | ForEach-Object {
        [pscustomobject]@{ Name = [string]$_.FriendlyName; Class = [string]$_.Class; Code = (Get-ErrorCode $_); InstanceId = [string]$_.InstanceId }
    })

    $touchRaw = @(Find-TouchDevices -Devices $all)

    $touch = @($touchRaw | ForEach-Object {
        $isHere = Test-DevicePresent $_
        [pscustomobject]@{
            Name         = [string]$_.FriendlyName
            InstanceId   = [string]$_.InstanceId
            Present      = $isHere
            ErrorCode    = if ($isHere) { Get-ErrorCode $_ } else { 0 }
            FirstInstall = Get-DeviceDate $_.InstanceId 'DEVPKEY_Device_FirstInstallDate'
            InstallDate  = Get-DeviceDate $_.InstanceId 'DEVPKEY_Device_InstallDate'
            LastArrival  = Get-DeviceDate $_.InstanceId 'DEVPKEY_Device_LastArrivalDate'
            LastRemoval  = Get-DeviceDate $_.InstanceId 'DEVPKEY_Device_LastRemovalDate'
        }
    })

    $controllers = @($all | Where-Object {
        (Test-DevicePresent $_) -and $_.Class -ne 'HIDClass' -and
        "$($_.FriendlyName)" -match 'I2C|Serial IO|Serial I/O|Touch Host Controller|\bTHC\b'
    } | ForEach-Object {
        [pscustomobject]@{ Name = [string]$_.FriendlyName; Code = (Get-ErrorCode $_) }
    })

    Write-Step 'Reading screen information...'
    $screens = @(Get-ScreenList -Devices $all)

    [pscustomobject]@{
        Today           = Get-Date
        Manufacturer    = [string]$cs.Manufacturer
        Model           = [string]$friendly
        SystemModel     = [string]$cs.Model
        Sku             = [string]$cs.SystemSKUNumber
        Serial          = [string]$bios.SerialNumber
        Chassis         = [string]$bb.Product
        BoardMaker      = [string]$bb.Manufacturer
        CpuMaker        = [string]$cpu.Manufacturer
        BiosVersion     = [string]$bios.SMBIOSBIOSVersion
        BiosDate        = $bios.ReleaseDate
        OsCaption       = [string]$os.Caption
        Build           = [int]$os.BuildNumber
        Ubr             = [string]$cv.UBR
        DisplayVersion  = [string]$cv.DisplayVersion
        EditionId       = [string]$cv.EditionID
        OriginalInstall = Get-OriginalInstallDate
        LastUpdate      = if ($lastFix) { $lastFix.InstalledOn } else { $null }
        AppNames        = $apps
        WindowsUpdate   = $wu
        Problems        = $problems
        Touch           = $touch
        Controllers     = $controllers
        Screens         = $screens
    }
}

# =====================================================================
#  REPORT
# =====================================================================
$script:ReportLines = New-Object System.Collections.Generic.List[object]
$script:Actions     = New-Object System.Collections.Generic.List[string]

function Add-Line {
    param([string]$Text = '', [string]$Kind = 'info')
    $script:ReportLines.Add([pscustomobject]@{ Text = $Text; Kind = $Kind })
}
function Add-Head {
    param([string]$Title)
    Add-Line ''
    Add-Line ('=' * 66) 'rule'
    Add-Line "  $Title" 'head'
    Add-Line ('=' * 66) 'rule'
}
function Add-Field {
    param([string]$Name, $Value, [string]$Kind = 'info')
    $text = "$Value"
    if ($text.Trim() -eq '') { $text = '(not reported)' }
    Add-Line ('  {0,-24}: {1}' -f $Name, $text) $Kind
}
function Add-Action {
    param([string]$Text)
    $script:Actions.Add($Text)
}

function Get-WhereToDownload {
    # One line telling the customer where drivers / BIOS come from
    param($Brand, $Facts)
    $idValue = $Facts.Serial
    if ($Brand.IdHint -match 'model' -and $Brand.IdHint -notmatch 'serial') { $idValue = $Facts.SystemModel }
    if ($Brand.Drivers) { return "$($Brand.Drivers)  (enter your $($Brand.IdHint): $idValue)" }
    if ($Brand.Name -like 'Clevo*') { return "the company that sold this laptop (chassis code: $($Facts.Chassis))" }
    if ($Brand.Name -eq 'unknown manufacturer') { return 'the support website of the company that made this laptop' }
    return "the $($Brand.Name) support website"
}

function Build-Report {
    param($Facts)
    $f = $Facts
    $script:ReportLines.Clear()
    $script:Actions.Clear()

    $brand  = Get-BrandProfile -Manufacturer $f.Manufacturer -Model $f.SystemModel -BoardMaker $f.BoardMaker
    $hasApp = @($brand.Apps).Count -gt 0
    $found  = @()
    if ($hasApp) { $found = @(Find-SupportApp -Brand $brand -AppNames $f.AppNames) }
    $where  = Get-WhereToDownload -Brand $brand -Facts $f
    $updateTool = if ($found.Count -gt 0) { $found[0] } elseif ($hasApp) { $brand.Recommend } else { '' }

    Add-Line ''
    Add-Line '##################################################################' 'head'
    Add-Line '#                     LAPTOP CHECK REPORT                        #' 'head'
    Add-Line '#               Please email this file to support.               #' 'head'
    Add-Line '##################################################################' 'head'
    Add-Line ''
    $created = $f.Today.ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    Add-Line "  Created   : $created"
    if ($Label) { Add-Line "  Reference : $Label" }
    Add-Line "  Tool      : LaptopCheck v$ToolVersion (data reviewed $DataReviewed)"

    # ---------------------------------------------------------- 1
    Add-Head '1. LAPTOP'
    Add-Field 'Make'          $f.Manufacturer 'key'
    Add-Field 'Model'         $f.Model 'key'
    Add-Field 'System model'  $f.SystemModel
    Add-Field 'SKU'           $f.Sku
    Add-Field 'Serial number' $f.Serial 'key'
    if ($brand.Name -like 'Clevo*') { Add-Field 'Chassis code' $f.Chassis 'key' }

    # ---------------------------------------------------------- 2
    Add-Head "2. SUPPORT APP ($($brand.Name))"
    if (-not $hasApp) {
        Add-Line "  $($brand.NoAppNote)"
    }
    elseif ($found.Count -gt 0) {
        Add-Line "  [OK] Installed: $($found -join ', ')" 'ok'
        Add-Line "       $($brand.HowToUse)"
    }
    else {
        Add-Line "  [!] $($brand.Recommend) is NOT installed." 'act'
        Add-Line '      It installs the correct drivers and BIOS for this exact model.'
        Add-Line ''
        Add-Line '      How to install it:'
        foreach ($i in @($brand.Install)) { Add-Line "        $i" 'key' }
        Add-Action "Install $($brand.Recommend) ($(@($brand.Install)[0])), then use it to install all driver and BIOS updates."
    }

    # ---------------------------------------------------------- 3
    Add-Head '3. WINDOWS'
    $w = Get-WindowsVerdict -Build $f.Build -EditionId $f.EditionId -Today $f.Today
    $verText = if ($f.DisplayVersion) { $f.DisplayVersion } else { $w.Version }
    $buildText = if ($f.Ubr) { "$($f.Build).$($f.Ubr)" } else { "$($f.Build)" }
    Add-Field 'Windows' "$($f.OsCaption) $verText" 'key'
    Add-Field 'Build'   $buildText
    Add-Field 'Last Windows update' (Format-Date $f.LastUpdate)

    $howTo11 = "Settings > Windows Update > Check for updates. If a newer Windows 11 version is offered, choose 'Download and install'. If nothing is offered, use the Windows 11 Installation Assistant: $($Links.Win11Assistant)"

    switch ($w.State) {
        'Supported' {
            Add-Field 'Support' "OK - security updates until $(Format-Date $w.EndDate)" 'ok'
        }
        'EndingSoon' {
            Add-Field 'Support' "ENDS $(Format-Date $w.EndDate) - in $($w.DaysLeft) day(s)" 'act'
            Add-Action "Windows 11 $verText stops receiving security updates on $(Format-Date $w.EndDate). Update now: $howTo11"
        }
        'Ended' {
            Add-Field 'Support' "ENDED $(Format-Date $w.EndDate) - no security updates" 'bad'
            Add-Action "This Windows version no longer receives security updates. Update it: $howTo11"
        }
        'Newer' {
            Add-Field 'Support' 'OK - newer than this tool knows about' 'ok'
        }
        'LTSC' {
            Add-Field 'Support' 'Long-Term Servicing edition - check its own support date' 'info'
        }
        'Win10' {
            $esuEnd  = ConvertTo-Date $Win10EsuEnd
            $esuOpen = ($f.Today.Date -le $esuEnd)
            # An update newer than the Windows 10 end date means the PC is probably in the ESU programme
            $recent = ($f.LastUpdate -and
                       ([datetime]$f.LastUpdate) -gt (ConvertTo-Date $Win10EndOfSupport) -and
                       ([datetime]$f.LastUpdate) -gt $f.Today.AddDays(-75))
            if ($recent -and $esuOpen) {
                Add-Field 'Support' "Regular support ENDED $(Format-Date $w.EndDate). Extended Security Updates (ESU) appear to be active until $(Format-Date $esuEnd)" 'act'
                Add-Action "Windows 10 security updates stop on $(Format-Date $esuEnd) (ESU). Plan the move to Windows 11: check this laptop with PC Health Check ($($Links.PcHealthCheck)) and, if it qualifies, upgrade from Settings > Windows Update."
            }
            elseif ($esuOpen) {
                Add-Field 'Support' "Windows 10 support ENDED $(Format-Date $w.EndDate)" 'bad'
                Add-Action "Windows 10 no longer gets regular updates. Check if this laptop can run Windows 11 with PC Health Check ($($Links.PcHealthCheck)); if it can, upgrade from Settings > Windows Update. If it cannot, enrol in Extended Security Updates from Settings > Windows Update (consumer ESU runs until $(Format-Date $esuEnd))."
            }
            else {
                Add-Field 'Support' "Windows 10 support ENDED $(Format-Date $w.EndDate) (consumer ESU ended $(Format-Date $esuEnd))" 'bad'
                Add-Action "Windows 10 no longer gets any security updates. Check if this laptop can run Windows 11 with PC Health Check ($($Links.PcHealthCheck)) and upgrade from Settings > Windows Update."
            }
        }
        'Win10Old' {
            Add-Field 'Support' 'ENDED - this Windows 10 version gets no updates at all' 'bad'
            Add-Action "Update Windows 10 to version 22H2 (Settings > Windows Update), then check whether this laptop can move to Windows 11: $($Links.PcHealthCheck)"
        }
        default {
            Add-Field 'Support' 'Could not determine' 'info'
        }
    }

    $wu = $f.WindowsUpdate
    $updateActionAdded = $false
    if ($null -eq $wu) {
        Add-Line '  Windows Update was not checked (-SkipWindowsUpdate).'
    }
    elseif (-not $wu.Searched) {
        Add-Line '  [!] Windows Update could not be checked.' 'act'
        if ($wu.Error) { Add-Line "      Reason: $($wu.Error)" }
        Add-Action 'Open Settings > Windows Update and check for updates manually (this tool could not reach Windows Update).'
        $updateActionAdded = $true
    }
    else {
        $waiting = @($wu.Software) + @($wu.Feature)
        if ($waiting.Count -eq 0) {
            Add-Line '  [OK] No Windows updates are waiting.' 'ok'
        } else {
            Add-Line "  [!] $($waiting.Count) Windows update(s) waiting:" 'act'
            foreach ($t in $waiting) { Add-Line "        - $t" }
            Add-Action "Install the $($waiting.Count) waiting Windows update(s): Settings > Windows Update > Download and install, then restart."
            $updateActionAdded = $true
        }
    }
    $stale = ($f.LastUpdate -and ([datetime]$f.LastUpdate) -lt $f.Today.AddDays(-60))
    if ($stale -and $w.State -notmatch 'Win10') {
        Add-Line "  [!] The newest Windows update on record is from $(Format-Date $f.LastUpdate) (over 60 days ago)." 'act'
        if (-not $updateActionAdded) {
            Add-Action "Windows has not recorded an update since $(Format-Date $f.LastUpdate). Open Settings > Windows Update and install everything offered."
        }
    }

    # ---------------------------------------------------------- 4
    Add-Head '4. DRIVERS'
    if ($null -eq $wu) {
        Add-Line '  Driver updates from Windows Update were not checked.'
    }
    elseif ($wu.Searched) {
        $drv = @($wu.Drivers)
        if ($drv.Count -eq 0) {
            Add-Line '  [OK] Windows Update has no driver updates waiting.' 'ok'
        } else {
            Add-Line "  [!] $($drv.Count) driver update(s) available from Windows Update:" 'act'
            foreach ($d in $drv) {
                $tag = if ($d.Optional) { '  (optional)' } else { '' }
                Add-Line "        - $($d.Title)$tag"
            }
            Add-Action "Install the driver updates: Settings > Windows Update > Advanced options > Optional updates > Driver updates."
        }
    }

    $problems  = @($f.Problems)
    $touchIds  = @($f.Touch | ForEach-Object { $_.InstanceId })
    $nonTouch  = @($problems | Where-Object { $touchIds -notcontains $_.InstanceId })
    $noDriver  = @($nonTouch | Where-Object { $_.Code -eq 28 })
    $disabled  = @($nonTouch | Where-Object { $_.Code -eq 22 })
    $faulty    = @($nonTouch | Where-Object { $_.Code -ne 28 -and $_.Code -ne 22 })

    Add-Line ''
    if ($problems.Count -eq 0) {
        Add-Line '  [OK] No devices are reporting a problem.' 'ok'
    } else {
        Add-Line "  Devices reporting a problem: $($problems.Count)" 'act'
        foreach ($p in $problems) {
            $n = if ($p.Name) { $p.Name } else { "(unnamed device $(Get-ShortId $p.InstanceId))" }
            Add-Line "        - $n  [$(Get-ProblemText $p.Code)]"
        }
    }
    if ($noDriver.Count -gt 0) {
        Add-Action "$($noDriver.Count) device(s) have no driver installed. This usually means the chipset drivers are missing (common after Windows is reset). Install them with $(if ($updateTool) { $updateTool } else { 'the manufacturer drivers' }) or from $where."
    }
    if ($faulty.Count -gt 0) {
        Add-Action "$($faulty.Count) device(s) report a fault (see section 4). Reinstall their drivers from $where."
    }
    if ($disabled.Count -gt 0) {
        Add-Line "  Note: $($disabled.Count) device(s) are disabled. If that was not done on purpose, enable them in Device Manager."
    }

    Add-Line ''
    Add-Line '  Where to download drivers:' 'key'
    if ($updateTool) { Add-Line "        $updateTool (recommended - it picks the right drivers)" }
    Add-Line "        $where"
    if (-not $brand.Drivers -and -not $hasApp) {
        if ($f.CpuMaker -match 'Intel') { Add-Line "        Intel Driver & Support Assistant: $($Links.IntelDsa)" }
        if ($f.CpuMaker -match 'AMD')   { Add-Line "        AMD drivers: $($Links.AmdDrivers)" }
        Add-Line '        (Prefer the laptop maker''s drivers where they exist - generic chip'
        Add-Line '         drivers can replace settings the laptop maker customised.)'
    }

    # ---------------------------------------------------------- 5
    Add-Head '5. BIOS'
    Add-Field 'Installed version' $f.BiosVersion 'key'
    Add-Field 'Installed date'    (Format-Date $f.BiosDate)

    $fw = @()
    if ($wu -and $wu.Searched) { $fw = @($wu.Firmware) }
    if ($fw.Count -gt 0) {
        Add-Line ''
        Add-Line '  [!] A BIOS / firmware update is offered by Windows Update:' 'act'
        foreach ($d in $fw) { Add-Line "        - $($d.Title)" }
        Add-Action 'Install the BIOS / firmware update: Settings > Windows Update > Advanced options > Optional updates. Keep the charger connected and do not turn the laptop off while it installs.'
    }
    elseif ($wu -and $wu.Searched) {
        Add-Line ''
        Add-Line '  Windows Update does not offer a newer BIOS for this laptop.'
        Add-Line '  Not every manufacturer publishes BIOS updates there, so also check below.'
    }
    if ($f.BiosDate -and ([datetime]$f.BiosDate) -lt $f.Today.AddYears(-3)) {
        Add-Line ''
        Add-Line "  This BIOS is from $(([datetime]$f.BiosDate).Year). Older laptops often have no newer BIOS -"
        Add-Line '  update only if the manufacturer offers a newer version.'
    }
    Add-Line ''
    Add-Line '  Where to download the BIOS:' 'key'
    if ($updateTool) { Add-Line "        $updateTool (recommended)" }
    Add-Line "        $where"
    Add-Line '  Always keep the charger connected during a BIOS update.'

    # ---------------------------------------------------------- 6
    Add-Head '6. TOUCHSCREEN'
    # A touchscreen connects through an I2C HID device. If that device is faulty the
    # touchscreen itself may not show up at all, so it must not read as "never had touch".
    $hidFaults = @($nonTouch | Where-Object { $_.Class -eq 'HIDClass' -and "$($_.Name)" -match 'I2C|touch|digitizer|^$' })
    $missingDrivers = ($noDriver.Count -gt 0) -or ($hidFaults.Count -gt 0)
    $t = Get-TouchVerdict -TouchDevices $f.Touch -OriginalInstall $f.OriginalInstall -DriversMissing $missingDrivers

    Add-Field 'Windows first set up' (Format-Date $f.OriginalInstall)
    switch ($t.History) {
        'Yes'        { Add-Field 'Touch at first setup' 'YES - this laptop came with a touchscreen' 'ok' }
        'AddedLater' { Add-Field 'Touch at first setup' "NOT CONFIRMED - oldest touchscreen record is from $(Format-Date $t.FirstSeen)" 'act' }
        'NoDate'     { Add-Field 'Touch at first setup' 'UNKNOWN - a touchscreen is installed, but Windows does not say when it was added' 'info' }
        'CannotTell' {
            $why = if ($noDriver.Count -gt 0) { 'drivers are missing' } else { 'a touch (I2C HID) device is faulty' }
            Add-Field 'Touch at first setup' "CANNOT TELL - $why" 'act'
        }
        default      { Add-Field 'Touch at first setup' 'NO - no touchscreen has ever been detected' 'key' }
    }
    switch ($t.Now) {
        'Working'      { Add-Field 'Touchscreen now' 'WORKING' 'ok' }
        'Disabled'     { Add-Field 'Touchscreen now' 'DISABLED in Windows' 'bad' }
        'Error'        { Add-Field 'Touchscreen now' "ERROR ($(Get-ProblemText $t.Code))" 'bad' }
        'NotConnected' { Add-Field 'Touchscreen now' "NOT DETECTED - last seen $(Format-Date $t.LastSeen)" 'bad' }
        default        { Add-Field 'Touchscreen now' 'none' }
    }

    $ctrl = @($f.Controllers)
    if ($ctrl.Count -gt 0) {
        $okCtrl  = @($ctrl | Where-Object { $_.Code -eq 0 }).Count
        Add-Field 'Touch controllers' "$okCtrl of $($ctrl.Count) working" $(if ($okCtrl -eq $ctrl.Count) { 'ok' } else { 'bad' })
    } else {
        # Red only when touch is missing or broken: touch that connects over USB has no I2C controller
        Add-Field 'Touch controllers' 'none found (I2C / Serial IO / Touch Host Controller)' $(if ($t.History -ne 'Never' -and $t.Now -ne 'Working') { 'bad' } else { 'info' })
    }
    Add-Line '  (If Windows was reset or reinstalled, "first set up" means that date.)'
    if ($t.History -eq 'AddedLater') {
        Add-Line '  The oldest touchscreen record is newer than Windows. Either a touch screen was'
        Add-Line '  fitted later, or the older record was removed (driver uninstalled, Windows reset).'
        Add-Line '  If you know the original screen had touch, tell us when you send this report.'
    }

    $driverSource = if ($updateTool) { "$updateTool, or $where" } else { $where }
    $removeSteps = @(
        '1. Shut down fully: hold Shift while clicking Shut down. Wait 10 seconds, then turn it on.',
        '2. Remove the old touchscreen driver: open Device Manager > View > Show hidden devices.',
        '   Expand "Human Interface Devices". Right-click each "HID-compliant touch screen"',
        '   (including greyed-out ones) and choose Uninstall device.',
        '3. Click Action > Scan for hardware changes, then restart the laptop.',
        "4. Update the touch drivers - chipset, Serial IO (I2C) and Touch Host Controller - using $driverSource."
    )

    Add-Line ''
    switch ($t.Now) {
        'Working' {
            Add-Line '  [OK] Nothing to do.' 'ok'
        }
        'Disabled' {
            Add-Line '  What to do:' 'key'
            Add-Line '    Open Device Manager > Human Interface Devices. Right-click'
            Add-Line '    "HID-compliant touch screen" and choose Enable device.'
            Add-Action 'The touchscreen is switched off in Windows. Device Manager > Human Interface Devices > right-click "HID-compliant touch screen" > Enable device.'
        }
        { $_ -eq 'Error' -or $_ -eq 'NotConnected' } {
            if ($t.Now -eq 'NotConnected') {
                Add-Line '  If this laptop now has a screen WITHOUT touch, this is expected - nothing to fix.' 'info'
                Add-Line ''
            }
            Add-Line '  What to do - remove and update the touchscreen driver:' 'key'
            foreach ($s in $removeSteps) { Add-Line "    $s" }
            Add-Line '    5. If touch is still not detected, the screen''s touch cable may be loose or'
            Add-Line '       the touch layer faulty. Contact us and attach this report.'
            if ($t.Now -eq 'NotConnected') {
                Add-Action "The touchscreen is not detected (last seen $(Format-Date $t.LastSeen)). If this screen should have touch, follow the steps in section 6: remove the touchscreen driver in Device Manager, restart, then update the chipset and Serial IO (I2C) drivers. If the screen has no touch layer, nothing needs fixing."
            } else {
                Add-Action 'The touchscreen is not working. Follow the steps in section 6: remove the touchscreen driver in Device Manager, restart, then update the chipset and Serial IO (I2C) drivers.'
            }
        }
        default {
            if ($t.History -eq 'CannotTell') {
                if ($hidFaults.Count -gt 0) {
                    Add-Line '  A touch (HID / I2C) device is reporting a fault. Touchscreens connect this way,' 'act'
                    Add-Line '  so the touchscreen may be hidden behind it:'
                    foreach ($h in $hidFaults) {
                        $hn = if ($h.Name) { $h.Name } else { "(unnamed device $(Get-ShortId $h.InstanceId))" }
                        Add-Line "        - $hn  [$(Get-ProblemText $h.Code)]"
                    }
                    Add-Line ''
                    Add-Line '  What to do:' 'key'
                    Add-Line '    1. Shut down fully: hold Shift while clicking Shut down. Wait 10 seconds, then turn it on.'
                    Add-Line '    2. Device Manager > View > Show hidden devices > Human Interface Devices.'
                    Add-Line '       Right-click each device listed above and choose Uninstall device.'
                    Add-Line '    3. Click Action > Scan for hardware changes, then restart the laptop.'
                    Add-Line "    4. Update the touch drivers - chipset, Serial IO (I2C) and Touch Host Controller - using $driverSource."
                    Add-Action 'A touch (I2C HID) device is faulty, which can hide the touchscreen. Follow the steps in section 6, then run this check again.'
                }
                if ($noDriver.Count -gt 0) {
                    Add-Line '  Drivers are missing, so a touchscreen could not be detected even if fitted.' 'act'
                    Add-Line "  Install the chipset drivers ($driverSource), restart, then run this check again."
                }
            } else {
                Add-Line '  No touchscreen is detected, and none is recorded in the Windows device history.'
                Add-Line '  Many models have a non-touch version with no touch connection on the motherboard:'
                Add-Line '  a touch screen fitted to it displays normally but touch will not work.'
                Add-Line '  If this laptop should have touch, restart it and run this check again, then'
                Add-Line '  contact us and attach this report.'
            }
        }
    }

    # ---------------------------------------------------------- 7
    Add-Head '7. SCREEN'
    $sel = Select-Screens -Screens $f.Screens
    if (-not $sel.Now) {
        Add-Line '  The laptop screen could not be read (is the lid closed?).' 'act'
    } else {
        $e = $sel.Now.Edid
        Add-Field 'PnP ID'      $e.PnpId 'key'
        Add-Field 'Part number' $(if ($e.Part) { $e.Part } else { 'not stored in this screen' }) 'key'
        Add-Field 'Resolution'  $(if ($e.Width) { "$($e.Width) x $($e.Height)  $($e.Aspect)" } else { '' })
        Add-Field 'Max refresh' $(if ($e.MaxHz) { "$($e.MaxHz) Hz" } else { '' })
        $sizeText = ''
        if ($e.WidthMm -gt 0)  { $sizeText = "$($e.WidthMm) x $($e.HeightMm) mm  (about $($e.Diagonal) inches)" }
        elseif ($e.Diagonal)   { $sizeText = "about $($e.Diagonal) inches" }
        Add-Field 'Size'        $sizeText
        Add-Field 'Made'        $e.Built
    }
    $earlier = @($sel.Earlier)
    Add-Line ''
    if ($earlier.Count -eq 0) {
        Add-Line '  No earlier laptop screens are recorded.'
    } else {
        Add-Line '  Earlier laptop-size screens recorded here (may include a portable monitor):' 'key'
        foreach ($s in $earlier) {
            $p = if ($s.Edid.Part) { "  $($s.Edid.Part)" } else { '' }
            $when = Format-Date $s.LastArrival
            Add-Line ('        {0}{1}  {2}x{3}  {4} Hz  - last connected {5}' -f $s.Edid.PnpId, $p, $s.Edid.Width, $s.Edid.Height, $s.Edid.MaxHz, $when)
        }
    }

    # ---------------------------------------------------------- 8
    Add-Head 'WHAT TO DO'
    if ($script:Actions.Count -eq 0) {
        Add-Line '  [OK] Nothing needs attention.' 'ok'
    } else {
        $n = 1
        foreach ($a in $script:Actions) {
            Add-Line "  $n. $a" 'act'
            Add-Line ''
            $n++
        }
    }
    Add-Line ''
    Add-Line '##################################################################' 'head'
    Add-Line '#                         END OF REPORT                          #' 'head'
    Add-Line '##################################################################' 'head'
}

function Show-Report {
    $colors = @{ head = 'Cyan'; rule = 'DarkGray'; ok = 'Green'; act = 'Yellow'; bad = 'Red'; key = 'White'; info = 'Gray' }
    foreach ($l in $script:ReportLines) {
        $c = $colors[$l.Kind]
        if (-not $c) { $c = 'Gray' }
        Write-Host $l.Text -ForegroundColor $c
    }
}

function Save-Report {
    $dir = $OutputPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        # A technician asked for this folder: create it rather than silently saving elsewhere
        try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch { }
    }
    if (-not $dir) { $dir = [Environment]::GetFolderPath('Desktop') }
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { $dir = $env:USERPROFILE }
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { $dir = $env:TEMP }

    $stamp = (Get-Date).ToString('yyyy-MM-dd_HHmm', [Globalization.CultureInfo]::InvariantCulture)
    $name  = "LaptopCheck_$stamp.txt"
    $file = Join-Path $dir $name
    $text = ($script:ReportLines | ForEach-Object { $_.Text }) -join "`r`n"
    try {
        Set-Content -LiteralPath $file -Value $text -Encoding UTF8 -ErrorAction Stop
    } catch {
        $file = Join-Path $env:TEMP $name
        Set-Content -LiteralPath $file -Value $text -Encoding UTF8
    }
    return $file
}

function Invoke-LaptopCheck {
    Clear-Host
    Write-Host ''
    Write-Host '  Checking this laptop. This can take a few minutes.' -ForegroundColor Cyan
    Write-Host '  Nothing on this computer will be changed.' -ForegroundColor Gray
    Write-Host ''

    $facts = Get-AllFacts
    Build-Report -Facts $facts

    Clear-Host
    Show-Report
    $file = Save-Report

    Write-Host ''
    Write-Host '  ----------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host '   Report saved to:' -ForegroundColor Green
    Write-Host "   $file" -ForegroundColor White
    Write-Host '   Please email this file to support.' -ForegroundColor Gray
    Write-Host '  ----------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$file`"" -ErrorAction SilentlyContinue
    if (-not $NoPause) { Read-Host '  Press Enter to close' | Out-Null }
}

# Run only when started as a script (not when loaded for testing)
if ($MyInvocation.InvocationName -ne '.') { Invoke-LaptopCheck }
