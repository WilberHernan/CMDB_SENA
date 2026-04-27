<#
.SYNOPSIS
  Script de inventario de hardware para CMDB SENA CCYS.
  Lee especificaciones del equipo local y abre el navegador en el CMDB.

.DESCRIPTION
  Ejecutar en el equipo nuevo (Windows).
  Detecta hardware via WMI, pregunta por la placa SENA y abre la URL del
  deployment web de Apps Script con query parameters listos para autorellenar.

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
$script:CMDB_URL = "https://script.google.com/macros/s/AKfycby7Z_AebjQMmbXSML_VwKFVY0XFi5HPMJSzCYz0RFzqkMi28xY7bI61BpB6bJbI-h0/exec"

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
    return "$GB GB"
}

function Get-MemoryTypeString {
    try {
        $mem = Get-WmiObject Win32_PhysicalMemory -ErrorAction Stop | Select-Object -First 1
        if (-not $mem) { return "" }
        switch ($mem.SMBIOSMemoryType) {
            24 { return "DDR3" }
            26 { return "DDR4" }
            34 { return "DDR5" }
            default { return "" }
        }
    } catch { return "" }
}

function Get-OSDisplayVersion {
    try {
        $reg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
        if ($reg.DisplayVersion) { return $reg.DisplayVersion }
        if ($reg.ReleaseId)      { return $reg.ReleaseId }
        return $reg.BuildNumber
    } catch { return "" }
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
# LECTURA DE HARDWARE
# ============================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  INVENTARIO CMDB - SENA CCYS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Detectando hardware del equipo..." -ForegroundColor DarkGray
Write-Host ""

$comp   = $null
$bios   = $null
$cpu    = $null
$os     = $null
$vid    = $null
$macEth  = ""
$macWifi = ""

try { $comp = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop } catch {}
try { $bios = Get-WmiObject Win32_BIOS -ErrorAction Stop } catch {}
try { $cpu  = Get-WmiObject Win32_Processor -ErrorAction Stop | Select-Object -First 1 } catch {}
try { $os   = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop } catch {}
try { $vid  = Get-WmiObject Win32_VideoController -ErrorAction Stop | Select-Object -First 1 } catch {}

# DETECCION DE DISCOS
Write-Host ""
Write-Host "Detectando discos..." -ForegroundColor DarkGray

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
            Write-Host "  [OK] Disco $idx`: $tipo / $tam (Modelo: $model)" -ForegroundColor DarkGray
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
            Write-Host "  [OK] Disco C: detectado por emergencia: HDD / $tam" -ForegroundColor DarkGray
        }
    } catch {}
}

# DETECCION DE RED
try {
    $net = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' -or $_.HardwareInterface -eq $true }
    $macEth  = Format-MacAddress (($net | Where-Object { $_.PhysicalMediaType -eq '802.3' -or $_.Name -match 'Ethernet' } | Select-Object -First 1).MacAddress)
    $macWifi = Format-MacAddress (($net | Where-Object { $_.PhysicalMediaType -match '802.11' -or $_.Name -match 'Wi-?Fi|Wireless' } | Select-Object -First 1).MacAddress)
    Write-Host "  [OK] MAC Cableada: $macEth | MAC WiFi: $macWifi" -ForegroundColor DarkGray
} catch {
    try {
        $wmiNet = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop | Where-Object { $_.IPEnabled -eq $true }
        foreach ($adapter in $wmiNet) {
            $mac = Format-MacAddress $adapter.MACAddress
            if (-not $macEth -and $adapter.DefaultIPGateway) { $macEth = $mac }
            elseif (-not $macWifi) { $macWifi = $mac }
        }
        Write-Host "  [OK] MAC detectada via WMI: $macEth / $macWifi" -ForegroundColor DarkGray
    } catch {}
}

# Fecha actual para registros nuevos
$fechaHoy = Get-Date -Format "dd/MM/yyyy"

# Armar diccionario de parametros
$params = @{
    modo                = "nuevo"
    placa               = ""
    hostname            = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { "" }
    marca               = if ($comp -and $comp.Manufacturer) { $comp.Manufacturer } else { "" }
    modelo              = if ($comp -and $comp.Model) { $comp.Model } else { "" }
    serial              = if ($bios -and $bios.SerialNumber) { $bios.SerialNumber } else { "" }
    procesador          = if ($cpu -and $cpu.Name) { $cpu.Name.Trim() } else { "" }
    ram                 = if ($comp -and $comp.TotalPhysicalMemory) { Get-RamStandard -Bytes $comp.TotalPhysicalMemory } else { "" }
    tipo_memoria        = Get-MemoryTypeString
    video               = if ($vid -and $vid.Name) { $vid.Name.Trim() } else { "" }
    so                  = if ($os -and $os.Caption) { $os.Caption.Trim() } else { "" }
    version_so          = Get-OSDisplayVersion
    mac_cableada        = $macEth
    mac_wifi            = $macWifi
    fecha_mantenimiento = $fechaHoy
    fecha_impacto       = $fechaHoy
}

# Asignar valores de discos detectados
$params.disco1_tipo = $disk1_tipo
$params.disco1_tam  = Normalize-DiskSize -Raw $disk1_tam
$params.disco2_tipo = $disk2_tipo
$params.disco2_tam  = Normalize-DiskSize -Raw $disk2_tam

# ============================================
# ENTRADA DE PLACA
# ============================================
Write-Host ""
Write-Host "Escanea la placa SENA del equipo y presiona ENTER: " -ForegroundColor Yellow -NoNewline
$placaInput = Read-Host
$placaInput = $placaInput.Trim().ToUpper() -replace "[^A-Z0-9\-]", ""

if (-not $placaInput) {
    Write-Host "ERROR: La placa es obligatoria. Proceso cancelado." -ForegroundColor Red
    Write-Host "Presiona cualquier tecla para cerrar..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
$params.placa = $placaInput

# ============================================
# PREVIEW DE DATOS
# ============================================
Write-Host ""
Write-Host "Datos detectados:" -ForegroundColor Green
Write-Host "  Placa:        $($params.placa)"
Write-Host "  Hostname:     $($params.hostname)"
Write-Host "  Marca/Modelo: $($params.marca) / $($params.modelo)"
Write-Host "  Serial:       $($params.serial)"
Write-Host "  CPU:          $($params.procesador)"
Write-Host "  RAM:          $($params.ram)"
Write-Host "  Disco 1:      $($params.disco1_tipo) / $($params.disco1_tam)"
if ($diskCount -gt 1) {
    Write-Host "  Disco 2:      $($params.disco2_tipo) / $($params.disco2_tam)"
}
Write-Host "  SO:           $($params.so) $($params.version_so)"
Write-Host ""

# ============================================
# CONSTRUIR URL Y ABRIR NAVEGADOR
# ============================================
$qsParts = @()
foreach ($kv in $params.GetEnumerator()) {
    $encoded = Url-Encode -Value $kv.Value
    $qsParts += "$($kv.Key)=$encoded"
}
$finalUrl = "$CMDB_URL`?$($qsParts -join '&')"

Write-Host "Abriendo CMDB en el navegador..." -ForegroundColor Green
try {
    Start-Process $finalUrl
} catch {
    Write-Host "ERROR al abrir el navegador: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Copia la URL de abajo manualmente:" -ForegroundColor Yellow
    Write-Host $finalUrl -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Listo. Completa los datos faltantes en la pagina y guarda." -ForegroundColor Cyan
Write-Host "Presiona cualquier tecla para cerrar esta ventana..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
