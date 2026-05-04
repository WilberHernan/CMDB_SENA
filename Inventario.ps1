<#
.SYNOPSIS
  Script de inventario de hardware UNIVERSAL para CMDB SENA CCYS.
  Lee especificaciones del equipo local con multiples metodos de fallback.

.DESCRIPTION
  Ejecutar en el equipo nuevo (Windows).
  Detecta hardware via WMI/CIM/Registry/CLI con fallback por capas.
  Abre el navegador en el CMDB con query parameters listos para autorellenar.

.USAGE
  1. Copiar este archivo al equipo nuevo.
  2. Click derecho > "Ejecutar con PowerShell"
  3. Escanear o escribir la placa cuando lo solicite.
  4. El navegador se abre solo con los datos detectados.

.NOTES
  Requiere configurar la URL del deployment de Apps Script en $CMDB_URL.
#>

# ============================================
# CONFIGURACION OBLIGATORIA
# ============================================
$script:CMDB_URL = "https://script.google.com/macros/s/AKfycbwExk7uXIjFyVKF941ZA1XMyDpVG5dOD8FAQkmLQ6AdAejhtDV4zvf-m_qP05aToYUc/exec"

if ($CMDB_URL -match "XXXXXXXX") {
    Write-Host ""
    Write-Host "ERROR: Debes configurar la URL del CMDB." -ForegroundColor Red
    Write-Host "Abre este archivo en un editor de texto y reemplaza la variable `$CMDB_URL" -ForegroundColor Yellow
    Write-Host "con la URL de tu Web App de Apps Script." -ForegroundColor Yellow
    Write-Host ""
    pause
    exit 1
}

# ============================================
# FUNCIONES AUXILIARES
# ============================================

function Get-RoundedDiskSize {
    param([long]$Bytes)
    if ($Bytes -le 0) { return "" }
    $GB = [math]::Round($Bytes / 1GB)
    if ($GB -lt 150)  { return "120 GB" }
    if ($GB -lt 380)  { return "256 GB" }
    if ($GB -lt 750)  { return "512 GB" }
    if ($GB -lt 1500) { return "1 TB" }
    if ($GB -lt 2500) { return "2 TB" }
    if ($GB -lt 4500) { return "4 TB" }
    return "$GB GB"
}

function Normalize-DiskSize {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return "" }
    $clean = $Raw.Trim().ToUpper() -replace "\s+", " "
    if ($clean -match "^(\d+)\s*(GB|TB)$") {
        $num = $Matches[1]
        $unit = $Matches[2]
        return "$num $unit"
    }
    return $clean
}

function Get-RamStandard {
    param([long]$Bytes)
    if ($Bytes -le 0) { return "" }
    $GB = [math]::Round($Bytes / 1GB)
    if ($GB -le 5)  { return "4 GB" }
    if ($GB -le 10) { return "8 GB" }
    if ($GB -le 18) { return "16 GB" }
    if ($GB -le 36) { return "32 GB" }
    if ($GB -le 72) { return "64 GB" }
    if ($GB -le 140) { return "128 GB" }
    return "$GB GB"
}

function Get-MemoryTypeString {
    # Metodo 1: SMBIOSMemoryType de Win32_PhysicalMemory
    try {
        $mem = Get-WmiObject Win32_PhysicalMemory -ErrorAction Stop | Select-Object -First 1
        if ($mem -and $mem.SMBIOSMemoryType) {
            switch ($mem.SMBIOSMemoryType) {
                20 { return "DDR" }
                21 { return "DDR2" }
                24 { return "DDR3" }
                26 { return "DDR4" }
                34 { return "DDR5" }
            }
        }
    } catch {}

    # Metodo 2: MemoryType (legacy, menos preciso)
    try {
        $mem = Get-WmiObject Win32_PhysicalMemory -ErrorAction Stop | Select-Object -First 1
        if ($mem -and $mem.MemoryType) {
            switch ($mem.MemoryType) {
                0  { return "DDR" }
                1  { return "DDR" }
                2  { return "DDR2" }
                3  { return "DDR2" }
                4  { return "DDR3" }
                5  { return "DDR3" }
                6  { return "DDR4" }
                7  { return "DDR4" }
            }
        }
    } catch {}

    # Metodo 3: CIM_PhysicalMemory
    try {
        $cimMem = Get-CimInstance CIM_PhysicalMemory -ErrorAction Stop | Select-Object -First 1
        if ($cimMem -and $cimMem.SMBIOSMemoryType) {
            switch ($cimMem.SMBIOSMemoryType) {
                20 { return "DDR" }
                21 { return "DDR2" }
                24 { return "DDR3" }
                26 { return "DDR4" }
                34 { return "DDR5" }
            }
        }
    } catch {}

    return ""
}

function Get-OSDisplayVersion {
    # Metodo 1: Registry DisplayVersion / ReleaseId
    try {
        $reg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
        if ($reg.DisplayVersion) { return $reg.DisplayVersion }
        if ($reg.ReleaseId)      { return $reg.ReleaseId }
        if ($reg.BuildNumber)    { return $reg.BuildNumber }
    } catch {}

    # Metodo 2: Win32_OperatingSystem Version
    try {
        $os = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
        if ($os -and $os.Version) { return $os.Version }
    } catch {}

    # Metodo 3: ver command
    try {
        $ver = [System.Environment]::OSVersion.Version
        return "$($ver.Major).$($ver.Minor).$($ver.Build)"
    } catch {}

    return ""
}

function Format-MacAddress {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return "" }
    return ($Raw -replace "-", ":" -replace " ", "").ToUpper()
}

function Url-Encode {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    return [System.Uri]::EscapeDataString($Value)
}

# ============================================
# DETECCION ROBUSTA DE HOSTNAME
# ============================================
function Get-HostnameUniversal {
    # Metodo 1: Variable de entorno (mas rapido y confiable)
    if ($env:COMPUTERNAME) { return $env:COMPUTERNAME.Trim().ToUpper() }

    # Metodo 2: Environment .NET
    try { $hn = [System.Environment]::MachineName; if ($hn) { return $hn.Trim().ToUpper() } } catch {}

    # Metodo 3: DNS
    try { $hn = [System.Net.Dns]::GetHostName(); if ($hn) { return $hn.Trim().ToUpper() } } catch {}

    # Metodo 4: WMI ComputerSystem
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.Name) { return $cs.Name.Trim().ToUpper() }
    } catch {}

    # Metodo 5: CIM ComputerSystem
    try {
        $cs = Get-CimInstance CIM_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.Name) { return $cs.Name.Trim().ToUpper() }
    } catch {}

    # Metodo 6: hostname.exe
    try {
        $hn = (& hostname.exe 2>$null)
        if ($hn) { return $hn.Trim().ToUpper() }
    } catch {}

    return "DESCONOCIDO"
}

# ============================================
# DETECCION ROBUSTA DE RED
# ============================================
function Get-NetworkInfoUniversal {
    $macEth = ""
    $macWifi = ""

    # Metodo 1: Get-NetAdapter (PowerShell nativo, requiere Win8+/2012+)
    try {
        $adapters = Get-NetAdapter -ErrorAction Stop |
            Where-Object {
                ($_.Status -eq 'Up' -or $_.HardwareInterface -eq $true) -and
                $_.MacAddress -and ($_.MacAddress -match "..") -and
                $_.HardwareInterface -eq $true
            }

        foreach ($a in $adapters) {
            $mac = Format-MacAddress $a.MacAddress
            if (-not $mac -or $mac -eq "00:00:00:00:00:00") { continue }

            $isEth = ($a.PhysicalMediaType -eq '802.3') -or
                     ($a.Name -match 'Ethernet|Ether|LAN|GBE|Gigabit') -or
                     ($a.InterfaceDescription -match 'Realtek|Intel.*Ethernet|Broadcom.*Net|Marvell')
            $isWifi = ($a.PhysicalMediaType -match '802.11|Native 802.11') -or
                      ($a.Name -match 'Wi-?Fi|Wireless|WLAN|802\.11') -or
                      ($a.InterfaceDescription -match 'Wireless|WiFi|WLAN|802\.11')

            if ($isEth -and -not $macEth) { $macEth = $mac }
            if ($isWifi -and -not $macWifi) { $macWifi = $mac }
        }
        if ($macEth -or $macWifi) {
            return @{ Ethernet = $macEth; WiFi = $macWifi; Source = "Get-NetAdapter" }
        }
    } catch {}

    # Metodo 2: Win32_NetworkAdapterConfiguration (WMI, mas compatible)
    try {
        $wmiNet = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop |
            Where-Object { $_.IPEnabled -eq $true -and $_.MACAddress }

        foreach ($cfg in $wmiNet) {
            $mac = Format-MacAddress $cfg.MACAddress
            if (-not $mac -or $mac -eq "00:00:00:00:00:00") { continue }

            # Obtener descripcion del adaptador padre
            $adapter = Get-WmiObject Win32_NetworkAdapter -ErrorAction Stop |
                Where-Object { $_.DeviceID -eq $cfg.Index } | Select-Object -First 1

            $desc = if ($adapter) { $adapter.Name + " " + $adapter.NetConnectionID } else { "" }
            $desc = $desc.ToUpper()

            $isEth = $desc -match 'ETHERNET|ETHER|LAN|GBE|GIGABIT|REALTEK.*PCI|INTEL.*ETHERNET|BROADCOM.*NET'
            $isWifi = $desc -match 'WIRELESS|WIFI|WLAN|802\.11|INTEL.*WIFI|REALTEK.*WIRELESS|QUALCOMM.*WIRELESS'

            if ($isEth -and -not $macEth) { $macEth = $mac }
            if ($isWifi -and -not $macWifi) { $macWifi = $mac }
        }
        if ($macEth -or $macWifi) {
            return @{ Ethernet = $macEth; WiFi = $macWifi; Source = "Win32_NetworkAdapterConfiguration" }
        }
    } catch {}

    # Metodo 3: getmac.exe (CLI nativo de Windows, siempre disponible)
    try {
        $getmac = (& getmac.exe /v /fo csv 2>$null) | ConvertFrom-Csv -ErrorAction Stop
        foreach ($entry in $getmac) {
            $mac = Format-MacAddress $entry.'Physical Address'
            if (-not $mac -or $mac -eq "00:00:00:00:00:00") { continue }

            $name = ($entry.'Connection Name' + " " + $entry.'Transport Name').ToUpper()
            $isEth = $name -match 'ETHERNET|ETHER|LAN|GBE|GIGABIT|PCI|NDIS'
            $isWifi = $name -match 'WIRELESS|WIFI|WLAN|802\.11'

            if ($isEth -and -not $macEth) { $macEth = $mac }
            if ($isWifi -and -not $macWifi) { $macWifi = $mac }
        }
        if ($macEth -or $macWifi) {
            return @{ Ethernet = $macEth; WiFi = $macWifi; Source = "getmac.exe" }
        }
    } catch {}

    # Metodo 4: ipconfig /all parsing
    try {
        $ipcfg = (& ipconfig.exe /all 2>$null)
        $currentDesc = ""
        foreach ($line in $ipcfg) {
            if ($line -match "^\s*Description[^:]*:\s*(.+)$") {
                $currentDesc = $Matches[1].Trim().ToUpper()
            }
            if ($line -match "^\s*Physical Address[^:]*:\s*([0-9A-F\-:]+)$") {
                $mac = Format-MacAddress $Matches[1]
                if (-not $mac -or $mac -eq "00:00:00:00:00:00") { continue }

                $isEth = $currentDesc -match 'ETHERNET|ETHER|LAN|GBE|GIGABIT|REALTEK.*PCI|INTEL.*ETHERNET'
                $isWifi = $currentDesc -match 'WIRELESS|WIFI|WLAN|802\.11|INTEL.*WIFI|REALTEK.*WIRELESS'

                if ($isEth -and -not $macEth) { $macEth = $mac }
                if ($isWifi -and -not $macWifi) { $macWifi = $mac }
            }
        }
        if ($macEth -or $macWifi) {
            return @{ Ethernet = $macEth; WiFi = $macWifi; Source = "ipconfig" }
        }
    } catch {}

    # Metodo 5: Win32_NetworkAdapter (sin filtro IPEnabled, para capturar desconectados tambien)
    try {
        $allAdapters = Get-WmiObject Win32_NetworkAdapter -ErrorAction Stop |
            Where-Object { $_.MACAddress -and $_.PhysicalAdapter -eq $true }

        foreach ($a in $allAdapters) {
            $mac = Format-MacAddress $a.MACAddress
            if (-not $mac -or $mac -eq "00:00:00:00:00:00") { continue }

            $name = ($a.Name + " " + $a.NetConnectionID).ToUpper()
            $isEth = $name -match 'ETHERNET|ETHER|LAN|GBE|GIGABIT'
            $isWifi = $name -match 'WIRELESS|WIFI|WLAN|802\.11'

            if ($isEth -and -not $macEth) { $macEth = $mac }
            if ($isWifi -and -not $macWifi) { $macWifi = $mac }
        }
        return @{ Ethernet = $macEth; WiFi = $macWifi; Source = "Win32_NetworkAdapter" }
    } catch {}

    return @{ Ethernet = $macEth; WiFi = $macWifi; Source = "none" }
}

# ============================================
# DETECCION ROBUSTA DE RAM TOTAL
# ============================================
function Get-RamTotalUniversal {
    # Metodo 1: Win32_ComputerSystem
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.TotalPhysicalMemory -and $cs.TotalPhysicalMemory -gt 0) {
            return [long]$cs.TotalPhysicalMemory
        }
    } catch {}

    # Metodo 2: Sumar Win32_PhysicalMemory
    try {
        $mems = Get-WmiObject Win32_PhysicalMemory -ErrorAction Stop
        $total = ($mems | Measure-Object -Property Capacity -Sum).Sum
        if ($total -and $total -gt 0) { return [long]$total }
    } catch {}

    # Metodo 3: CIM_PhysicalMemory
    try {
        $cimMems = Get-CimInstance CIM_PhysicalMemory -ErrorAction Stop
        $total = ($cimMems | Measure-Object -Property Capacity -Sum).Sum
        if ($total -and $total -gt 0) { return [long]$total }
    } catch {}

    # Metodo 4: systeminfo parsing
    try {
        $si = (& systeminfo.exe 2>$null) | Where-Object { $_ -match "Total Physical Memory" }
        if ($si -match "([\d,]+)\s*MB") {
            $mb = [long]($Matches[1] -replace ",", "")
            return $mb * 1MB
        }
        if ($si -match "([\d,]+)\s*GB") {
            $gb = [long]($Matches[1] -replace ",", "")
            return $gb * 1GB
        }
    } catch {}

    return 0
}

# ============================================
# DETECCION ROBUSTA DE CPU
# ============================================
function Get-CPUNameUniversal {
    # Metodo 1: Win32_Processor
    try {
        $cpu = Get-WmiObject Win32_Processor -ErrorAction Stop | Select-Object -First 1
        if ($cpu -and $cpu.Name) { return $cpu.Name.Trim() }
    } catch {}

    # Metodo 2: Registry (CPU name a veces esta aca)
    try {
        $reg = Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0" -ErrorAction Stop
        if ($reg.ProcessorNameString) { return $reg.ProcessorNameString.Trim() }
    } catch {}

    # Metodo 3: environment
    try {
        $proc = $env:PROCESSOR_IDENTIFIER
        if ($proc) { return $proc.Trim() }
    } catch {}

    return ""
}

# ============================================
# DETECCION ROBUSTA DE MARCA/MODELO
# ============================================
function Get-ComputerManufacturerUniversal {
    # Metodo 1: WMI
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.Manufacturer) { return $cs.Manufacturer.Trim() }
    } catch {}

    # Metodo 2: CIM
    try {
        $cs = Get-CimInstance CIM_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.Manufacturer) { return $cs.Manufacturer.Trim() }
    } catch {}

    # Metodo 3: Registry SystemInformation
    try {
        $reg = Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\BIOS" -ErrorAction Stop
        if ($reg.SystemManufacturer) { return $reg.SystemManufacturer.Trim() }
    } catch {}

    return ""
}

function Get-ComputerModelUniversal {
    # Metodo 1: WMI
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.Model) { return $cs.Model.Trim() }
    } catch {}

    # Metodo 2: CIM
    try {
        $cs = Get-CimInstance CIM_ComputerSystem -ErrorAction Stop
        if ($cs -and $cs.Model) { return $cs.Model.Trim() }
    } catch {}

    # Metodo 3: Registry
    try {
        $reg = Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\BIOS" -ErrorAction Stop
        if ($reg.SystemProductName) { return $reg.SystemProductName.Trim() }
    } catch {}

    return ""
}

# ============================================
# DETECCION ROBUSTA DE SERIAL
# ============================================
function Get-SerialUniversal {
    # Metodo 1: Win32_BIOS
    try {
        $bios = Get-WmiObject Win32_BIOS -ErrorAction Stop
        if ($bios -and $bios.SerialNumber -and $bios.SerialNumber -notmatch "^(To Be Filled|System Serial|Not Applicable|Default|1234567890|To be filled|SystemSerial|NotAvailable|NONE)$") {
            return $bios.SerialNumber.Trim()
        }
    } catch {}

    # Metodo 2: Win32_SystemEnclosure
    try {
        $se = Get-WmiObject Win32_SystemEnclosure -ErrorAction Stop
        if ($se -and $se.SerialNumber -and $se.SerialNumber -notmatch "^(To Be Filled|System Serial|Not Applicable|Default|1234567890)") {
            return $se.SerialNumber.Trim()
        }
    } catch {}

    # Metodo 3: Win32_ComputerSystemProduct
    try {
        $csp = Get-WmiObject Win32_ComputerSystemProduct -ErrorAction Stop
        if ($csp -and $csp.IdentifyingNumber -and $csp.IdentifyingNumber -notmatch "^(To Be Filled|System Serial|Not Applicable|Default|1234567890)") {
            return $csp.IdentifyingNumber.Trim()
        }
    } catch {}

    # Metodo 4: CIM_BIOSElement
    try {
        $cb = Get-CimInstance CIM_BIOSElement -ErrorAction Stop | Select-Object -First 1
        if ($cb -and $cb.SerialNumber -and $cb.SerialNumber -notmatch "^(To Be Filled|System Serial|Not Applicable|Default|1234567890)") {
            return $cb.SerialNumber.Trim()
        }
    } catch {}

    # Metodo 5: wmic bios get serialnumber
    try {
        $wmicSerial = (& wmic bios get serialnumber /value 2>$null) | Where-Object { $_ -match "SerialNumber=" }
        if ($wmicSerial -match "SerialNumber=(.+)$") {
            $s = $Matches[1].Trim()
            if ($s -and $s -notmatch "^(To Be Filled|System Serial|Not Applicable|Default|1234567890)") {
                return $s
            }
        }
    } catch {}

    return ""
}

# ============================================
# DETECCION ROBUSTA DE VIDEO
# ============================================
function Get-VideoControllerUniversal {
    # Metodo 1: Win32_VideoController (excluir genericos)
    try {
        $vids = Get-WmiObject Win32_VideoController -ErrorAction Stop |
            Where-Object {
                $_.Name -notmatch "Basic Display|Microsoft Basic|Standard VGA|VirtualBox|VMware|Hyper-V|Remote Desktop|Basic Render|Basic Display Adapter"
            } |
            Sort-Object AdapterRAM -Descending

        $best = $vids | Select-Object -First 1
        if ($best -and $best.Name) { return $best.Name.Trim() }
    } catch {}

    # Metodo 2: CIM_VideoController
    try {
        $vids = Get-CimInstance CIM_VideoController -ErrorAction Stop |
            Where-Object {
                $_.Name -notmatch "Basic Display|Microsoft Basic|Standard VGA|VirtualBox|VMware|Hyper-V|Remote Desktop"
            } |
            Sort-Object AdapterRAM -Descending
        $best = $vids | Select-Object -First 1
        if ($best -and $best.Name) { return $best.Name.Trim() }
    } catch {}

    # Metodo 3: Registry para GPU discreta
    try {
        $gpuKeys = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}" -ErrorAction Stop |
            Where-Object { $_.PSChildName -match "^\d{4}$" }
        foreach ($key in $gpuKeys) {
            $reg = Get-ItemProperty $key.PSPath -ErrorAction Stop
            if ($reg.DriverDesc -and ($reg.DriverDesc -notmatch "Basic Display|Microsoft Basic|Standard VGA")) {
                return $reg.DriverDesc.Trim()
            }
        }
    } catch {}

    return ""
}

# ============================================
# DETECCION ROBUSTA DE SO
# ============================================
function Get-OSCaptionUniversal {
    # Metodo 1: WMI
    try {
        $os = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
        if ($os -and $os.Caption) {
            return ($os.Caption.Trim() -replace "Microsoft ", "")
        }
    } catch {}

    # Metodo 2: CIM
    try {
        $os = Get-CimInstance CIM_OperatingSystem -ErrorAction Stop
        if ($os -and $os.Caption) {
            return ($os.Caption.Trim() -replace "Microsoft ", "")
        }
    } catch {}

    # Metodo 3: Registry
    try {
        $reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
        $prodName = $reg.ProductName
        if ($prodName) { return $prodName.Trim() }
    } catch {}

    return ""
}

# ============================================
# INICIO
# ============================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor DarkGray
Write-Host "  CMDB SENA CCYS - Inventario Universal" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Detectando hardware con multiples metodos..." -ForegroundColor DarkGray
Write-Host ""

# --- Hostname ---
$hostname = Get-HostnameUniversal
Write-Host "  Hostname:     $hostname" -ForegroundColor Green

# --- Marca / Modelo ---
$marca = Get-ComputerManufacturerUniversal
$modelo = Get-ComputerModelUniversal
Write-Host "  Marca:        $(if($marca){$marca}else{'No detectado'})" -ForegroundColor Green
Write-Host "  Modelo:       $(if($modelo){$modelo}else{'No detectado'})" -ForegroundColor Green

# --- Serial ---
$serial = Get-SerialUniversal
Write-Host "  Serial:       $(if($serial){$serial}else{'No detectado'})" -ForegroundColor Green

# --- CPU ---
$cpu = Get-CPUNameUniversal
Write-Host "  CPU:          $(if($cpu){$cpu}else{'No detectado'})" -ForegroundColor Green

# --- RAM ---
$ramBytes = Get-RamTotalUniversal
$ram = if ($ramBytes -gt 0) { Get-RamStandard -Bytes $ramBytes } else { "" }
Write-Host "  RAM:          $(if($ram){$ram}else{'No detectado'})" -ForegroundColor Green

# --- Tipo Memoria ---
$tipoMemoria = Get-MemoryTypeString
Write-Host "  Tipo Memoria: $(if($tipoMemoria){$tipoMemoria}else{'No detectado'})" -ForegroundColor Green

# --- Video ---
$video = Get-VideoControllerUniversal
Write-Host "  Video:        $(if($video){$video}else{'No detectado'})" -ForegroundColor Green

# --- SO ---
$so = Get-OSCaptionUniversal
$versionSO = Get-OSDisplayVersion
Write-Host "  S.O.:         $(if($so){$so}else{'No detectado'}) $(if($versionSO){$versionSO})" -ForegroundColor Green

# DETECCION DE DISCOS
Write-Host ""
Write-Host "-- Almacenamiento ------------------------------------------" -ForegroundColor DarkGray

$disk1_tipo = ""
$disk1_tam = ""
$disk2_tipo = "N/A"
$disk2_tam = "N/A"
$diskCount = 0

# METODO 1: Win32_DiskDrive
try {
    $dds = Get-WmiObject Win32_DiskDrive -ErrorAction Stop | Select-Object -First 2
    if ($dds) {
        if ($dds -isnot [System.Array]) { $dds = @($dds) }

        $idx = 0
        foreach ($d in $dds) {
            $idx++
            $tipo = "HDD"
            $model = if ($d.Model) { $d.Model.ToString().ToUpper().Trim() } else { "" }
            $interface = if ($d.InterfaceType) { $d.InterfaceType.ToString().ToUpper().Trim() } else { "" }
            $mediaType = if ($d.MediaType) { $d.MediaType.ToString().ToUpper().Trim() } else { "" }

            if ($mediaType -match "SSD|SOLID.STATE") { $tipo = "SSD" }
            elseif ($mediaType -match "NVME|EXTERNAL.HARD.DISK") { $tipo = "M2" }

            if ($interface -match "NVME|PCI") { $tipo = "M2" }
            elseif ($interface -match "USB|1394|SCSI") { $tipo = "HDD" }

            if ($model -match "NVME|NVMe|M\.2|M2\.0|PCIE.SSD|SM961|PM961|PM981|970.EVO|980.PRO|WD_BLACK.SN|SN730|SN750|SN850|SN570|WDC.*SN|SAMSUNG.*MZVL|SAMSUNG.*PM|INTEL.*SSD|ADATA.*SX") { $tipo = "M2" }
            elseif ($model -match "SSD|SOLID.STATE|SATA.SSD|CRUCIAL|KINGSTON.A|SAN_DISK|SANDISK|SAMSUNG.8|TEAM.*SSD|ADATA.*SP") { $tipo = "SSD" }
            elseif ($model -match "HDD|HARD.DISK|WDC.WD|ST\d|SEAGATE|HITACHI|TOSHIBA.MK|HGST") { $tipo = "HDD" }

            $rawSize = $d.Size
            $tam = Get-RoundedDiskSize -Bytes $rawSize

            if ($idx -eq 1) {
                $disk1_tipo = $tipo
                $disk1_tam = $tam
            } elseif ($idx -eq 2) {
                $disk2_tipo = $tipo
                $disk2_tam = $tam
            }
            $diskCount++
            Write-Host "  + Disco $idx`: $tipo / $tam" -ForegroundColor Green
        }
    }
} catch {}

# METODO 2: Emergencia (disco C:)
if ($diskCount -eq 0) {
    try {
        $ld = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
        if ($ld -and $ld.Size -gt 0) {
            $tamGB = [math]::Round($ld.Size / 1GB)
            $tam = ""
            if ($tamGB -lt 150)  { $tam = "120 GB" }
            elseif ($tamGB -lt 380)  { $tam = "256 GB" }
            elseif ($tamGB -lt 750)  { $tam = "512 GB" }
            elseif ($tamGB -lt 1500) { $tam = "1 TB" }
            else { $tam = "$tamGB GB" }
            $disk1_tipo = "HDD"
            $disk1_tam = $tam
            $diskCount = 1
            Write-Host "  + Disco C: HDD / $tam (emergencia)" -ForegroundColor Yellow
        }
    } catch {}
}

# --- Red ---
Write-Host ""
Write-Host "-- Red -----------------------------------------------------" -ForegroundColor DarkGray
$netInfo = Get-NetworkInfoUniversal
$macEth = $netInfo.Ethernet
$macWifi = $netInfo.WiFi
if ($macEth)  { Write-Host "  + MAC Cableada: $macEth [$($netInfo.Source)]" -ForegroundColor Green }
if ($macWifi) { Write-Host "  + MAC WiFi:     $macWifi [$($netInfo.Source)]" -ForegroundColor Green }
if (-not $macEth -and -not $macWifi) {
    Write-Host "  + No se detectaron adaptadores de red con MAC" -ForegroundColor Yellow
}

# --- Fecha ---
$fechaHoy = Get-Date -Format "dd/MM/yyyy"

# ============================================
# ARMA PARAMETROS
# ============================================
$params = @{
    modo                = "nuevo"
    placa               = ""
    hostname            = $hostname
    marca               = $marca
    modelo              = $modelo
    serial              = $serial
    procesador          = $cpu
    ram                 = $ram
    tipo_memoria        = $tipoMemoria
    video               = $video
    so                  = $so
    version_so          = $versionSO
    mac_cableada        = $macEth
    mac_wifi            = $macWifi
    fecha_mantenimiento = $fechaHoy
    fecha_impacto       = $fechaHoy
    disco1_tipo         = $disk1_tipo
    disco1_tam          = Normalize-DiskSize -Raw $disk1_tam
    disco2_tipo         = $disk2_tipo
    disco2_tam          = Normalize-DiskSize -Raw $disk2_tam
}

# ============================================
# PREVIEW DE DATOS
# ============================================
Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Datos detectados:" -ForegroundColor Green
Write-Host ""
Write-Host "  Hostname:     $($params.hostname)" -ForegroundColor White
Write-Host "  Marca:        $(if($params.marca){$params.marca}else{'---'})" -ForegroundColor White
Write-Host "  Modelo:       $(if($params.modelo){$params.modelo}else{'---'})" -ForegroundColor White
Write-Host "  Serial:       $(if($params.serial){$params.serial}else{'---'})" -ForegroundColor White
Write-Host "  CPU:          $(if($params.procesador){$params.procesador}else{'---'})" -ForegroundColor White
Write-Host "  RAM:          $(if($params.ram){$params.ram}else{'---'})" -ForegroundColor White
Write-Host "  Memoria:      $(if($params.tipo_memoria){$params.tipo_memoria}else{'---'})" -ForegroundColor White
Write-Host "  Disco 1:      $(if($params.disco1_tipo){$params.disco1_tipo}else{'---'}) / $(if($params.disco1_tam){$params.disco1_tam}else{'---'})" -ForegroundColor White
if ($diskCount -gt 1) {
    Write-Host "  Disco 2:      $($params.disco2_tipo) / $($params.disco2_tam)" -ForegroundColor White
}
Write-Host "  Video:        $(if($params.video){$params.video}else{'---'})" -ForegroundColor White
Write-Host "  S.O.:         $(if($params.so){$params.so}else{'---'}) $($params.version_so)" -ForegroundColor White
Write-Host "  MAC Eth:      $(if($params.mac_cableada){$params.mac_cableada}else{'---'})" -ForegroundColor White
Write-Host "  MAC WiFi:     $(if($params.mac_wifi){$params.mac_wifi}else{'---'})" -ForegroundColor White
Write-Host ""

# ============================================
# ENTRADA DE PLACA
# ============================================
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Escribe o escanea la placa y presiona ENTER: " -ForegroundColor White -NoNewline
$placaInput = Read-Host
$placaInput = $placaInput.Trim().ToUpper() -replace "'", "-" -replace "[^A-Z0-9\-]", ""

if (-not $placaInput) {
    Write-Host ""
    Write-Host "La placa es obligatoria. Proceso cancelado." -ForegroundColor Red
    Write-Host ""
    Write-Host "Presiona cualquier tecla para cerrar..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
$params.placa = $placaInput

# ============================================
# CONSTRUIR URL Y ABRIR NAVEGADOR
# ============================================
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

$qsParts = @()
foreach ($kv in $params.GetEnumerator()) {
    $encoded = Url-Encode -Value $kv.Value
    $qsParts += "$($kv.Key)=$encoded"
}
$finalUrl = "$CMDB_URL`?$($qsParts -join '&')"

Write-Host "Abriendo CMDB..." -ForegroundColor Green
Write-Host ""
try {
    Start-Process $finalUrl
    Write-Host "Navegador abierto." -ForegroundColor Green
} catch {
    Write-Host "No se pudo abrir el navegador." -ForegroundColor Red
    Write-Host "Copia esta URL:" -ForegroundColor Yellow
    Write-Host $finalUrl -ForegroundColor Yellow
}

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Completa los datos faltantes en la pagina y guarda." -ForegroundColor Gray
Write-Host ""
Write-Host "Presiona cualquier tecla para cerrar..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
