# === SMB/UNC Small Files Speed Test (Standalone PS 5.1) — v1.5.1 STRICT (per-run destination) ===
# - TXT por ejecución (junto al script)
# - CSV ÚNICO: 'smb_smallfiles_log.csv' con columnas: timestamp,run,files,total_MB,up_sec,up_mbps,down_sec,down_mbps,threads,cleanup_sec
# - Migración como v1.5 si detecta cabecera antigua.
# - NOVEDAD: $UniqueDestPerRun => subcarpeta de destino distinta por cada run (evita "skips").
#
# ---------- EDITA AQUI ----------
$RobocopyExe       = 'C:\Tools\robocopy.exe'
$LocalSourceDir  = 'C:\Velocidad\Archivos'     #  TU CARPETA LOCAL DE PRUEBA (muchos archivos)
$RemoteBase      = '\\172.0.1.100\SEPROTEC\L10N\Velocidad'
$Threads           = 16                          # /MT (prueba 8/16/32)
$Count             = 3                           # nº repeticiones por ejecución
$CleanupRemote     = $false                       # borrar carpeta remota al terminar
$DoUpload          = $true                       # medir subida
$DoDownload        = $true                       # medir bajada
$ExportCsv         = $true                       # escribir CSV
$UniqueDestPerRun  = $true                       # <<< crea \\...\run_<timestamp>\r1, r2, r3... (recomendado para runs realistas)
# ---- Opciones adicionales ----
$MaxFileKB         = $null                       # e.g. 256 (solo archivos <= 256 KB). $null = sin límite
$UseUnbuffered     = $false                      # true => añade /J (unbuffered); ojo con ficheros pequeños
# --------------------------------

$ErrorActionPreference = 'Stop'

# Salidas
$timestamp    = Get-Date -Format 'yyyyMMdd-HHmmss'
$ResultsPath  = Join-Path $PSScriptRoot ("smb_smallfiles_{0}.txt" -f $timestamp)  # TXT por ejecución
$CsvPath      = Join-Path $PSScriptRoot "smb_smallfiles_log.csv"                  # CSV único

# Helpers de escritura
function Append-TextLine { param([string]$Path,[string]$Text) $Text | Out-File -LiteralPath $Path -Encoding UTF8 -Append }
function Append-CsvLine  { param([string]$Path,[string]$CsvRow) $CsvRow | Out-File -LiteralPath $Path -Encoding UTF8 -Append }

# Cabeceras
$HeaderOld = "timestamp,run,files,total_MB,up_sec,up_mbps,down_sec,down_mbps"
$HeaderNew = "timestamp,run,files,total_MB,up_sec,up_mbps,down_sec,down_mbps,threads,cleanup_sec"

# Validaciones
if (!(Test-Path -LiteralPath $RobocopyExe)) { throw "Robocopy not found at $RobocopyExe" }
if (!(Test-Path -LiteralPath $LocalSourceDir)) { throw "Local source not found: $LocalSourceDir" }

# Encabezado TXT (por ejecución)
@("=== SMB/UNC Small Files Speed Test v1.5.1 (TXT + CSV con threads y cleanup) ===",
  "Local source: $LocalSourceDir",
  "Remote base:  $RemoteBase",
  "Robocopy:     $RobocopyExe",
  "Threads:      $Threads    Runs: $Count",
  ("Options:   MaxFileKB={0}  Unbuffered={1}  UniqueDestPerRun={2}" -f ($MaxFileKB), $UseUnbuffered, $UniqueDestPerRun),
  "Started:      $(Get-Date)",
  "------------------------------------------------------------") -join "`r`n" |
  Out-File -LiteralPath $ResultsPath -Encoding UTF8 -Force

# Preparar/Migrar CSV
if ($ExportCsv) {
  if (Test-Path -LiteralPath $CsvPath) {
    $first = (Get-Content -LiteralPath $CsvPath -TotalCount 1)
    if ($first -eq $HeaderOld) {
      $bak = "$CsvPath.bak_$timestamp"
      Rename-Item -LiteralPath $CsvPath -NewName $bak
      $HeaderNew | Out-File -LiteralPath $CsvPath -Encoding UTF8 -Force
      $oldLines = Get-Content -LiteralPath $bak | Select-Object -Skip 1
      foreach ($line in $oldLines) {
        if ($line -and $line.Trim().Length -gt 0) { Append-CsvLine -Path $CsvPath -CsvRow ("{0},," -f $line) }
      }
      Append-TextLine -Path $ResultsPath -Text ("[MIGRATION] Upgraded CSV header and carried over old rows from '{0}'" -f $bak)
    } elseif ($first -ne $HeaderNew) {
      $bak = "$CsvPath.bak_$timestamp"
      Rename-Item -LiteralPath $CsvPath -NewName $bak
      $HeaderNew | Out-File -LiteralPath $CsvPath -Encoding UTF8 -Force
      Append-TextLine -Path $ResultsPath -Text ("[MIGRATION] Unknown header. Started new CSV with header v1.5.1; old file saved as '{0}'" -f $bak)
    }
  } else {
    $HeaderNew | Out-File -LiteralPath $CsvPath -Encoding UTF8 -Force
  }
}

# Selección de archivos (tamaño máximo opcional)
$allFiles = Get-ChildItem -LiteralPath $LocalSourceDir -Recurse -File -Force
$files = if ($MaxFileKB -ne $null) {
  $limit = [int64]$MaxFileKB * 1024
  $allFiles | Where-Object { $_.Length -le $limit }
} else { $allFiles }

if (!$files -or $files.Count -eq 0) { throw "No files found in $LocalSourceDir (after size filter if applied)" }

$totalBytes = ($files | Measure-Object Length -Sum).Sum
$fileCount  = $files.Count
$mbTotal    = [math]::Round($totalBytes/1MB, 2)

# Args de Robocopy
function New-RobocopyArgs {
  param([string]$Src,[string]$Dst)
  $args = @($Src, $Dst, '/E','/R:0','/W:0','/NFL','/NDL','/NP', ("/MT:{0}" -f $Threads))
  if ($UseUnbuffered) { $args += '/J' }
  if ($MaxFileKB -ne $null) { $args += ("/MAX:{0}" -f ([int64]$MaxFileKB * 1024)) }
  return $args
}

# Medidor
function Measure-Copy {
  param([string]$Src,[string]$Dst)
  $args = New-RobocopyArgs -Src $Src -Dst $Dst
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  & $RobocopyExe @args | Out-Null
  $rc = $LASTEXITCODE
  $sw.Stop()
  if ($rc -ge 8) { throw ("Robocopy failed (RC={0})" -f $rc) }
  $secs = [math]::Round($sw.Elapsed.TotalSeconds, 2)
  $mbps = [math]::Round(($totalBytes*8)/$sw.Elapsed.TotalSeconds/1000000, 2)
  return @{ Seconds = $secs; Mbps = $mbps }
}

# Carpeta remota base por ejecución
$remoteRun = Join-Path $RemoteBase ("run_{0}" -f $timestamp)
New-Item -ItemType Directory -Force -Path $remoteRun | Out-Null

# Buffer de filas CSV para completar cleanup_sec al final
$csvRows = @()

for ($i=1; $i -le $Count; $i++) {
  Write-Host ""
  Write-Host ("[Run {0}/{1}] (MT={2})" -f $i, $Count, $Threads)

  # Destino específico por run (si está activado)
  $destThisRun = if ($UniqueDestPerRun) { Join-Path $remoteRun ("r{0}" -f $i) } else { $remoteRun }
  if ($DoUpload) { New-Item -ItemType Directory -Force -Path $destThisRun | Out-Null }

  $upSeconds = $null; $upMbps = $null
  $dnSeconds = $null; $dnMbps = $null

  if ($DoUpload) {
    Write-Host "  Upload folder..."
    $u = Measure-Copy -Src $LocalSourceDir -Dst $destThisRun
    $upSeconds = $u.Seconds
    $upMbps    = $u.Mbps
    $lineUp = "UPLOAD   {0} files / {1} MB in {2:n2}s  ->  {3} Mbps (MT={4})" -f $fileCount, $mbTotal, $upSeconds, $upMbps, $Threads
    Write-Host $lineUp
    Append-TextLine -Path $ResultsPath -Text $lineUp
  }

  if ($DoDownload) {
    Write-Host "  Download folder..."
    $localDl = Join-Path $PSScriptRoot ("smb_speedtest_temp\dl_{0}\r{1}" -f $timestamp, $i)
    if (Test-Path -LiteralPath $localDl) { Remove-Item -LiteralPath $localDl -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $localDl | Out-Null
    $srcForDownload = if ($DoUpload) { $destThisRun } else { $remoteRun }
    $d = Measure-Copy -Src $srcForDownload -Dst $localDl
    $dnSeconds = $d.Seconds
    $dnMbps    = $d.Mbps
    $lineDn = "DOWNLOAD {0} files / {1} MB in {2:n2}s  ->  {3} Mbps (MT={4})" -f $fileCount, $mbTotal, $dnSeconds, $dnMbps, $Threads
    Write-Host $lineDn
    Append-TextLine -Path $ResultsPath -Text $lineDn
  }

  # Guardar fila en buffer (cleanup_sec se completará al final)
  $now = Get-Date -Format 's'
  $row = New-Object PSObject -Property @{
    timestamp   = $now
    run         = $i
    files       = $fileCount
    total_MB    = $mbTotal
    up_sec      = $(if ($null -eq $upSeconds) { '' } else { [string]$upSeconds })
    up_mbps     = $(if ($null -eq $upMbps)    { '' } else { [string]$upMbps })
    down_sec    = $(if ($null -eq $dnSeconds) { '' } else { [string]$dnSeconds })
    down_mbps   = $(if ($null -eq $dnMbps)    { '' } else { [string]$dnMbps })
    threads     = $Threads
    cleanup_sec = ''   # se rellena más tarde
  }
  $csvRows += $row

  Append-TextLine -Path $ResultsPath -Text ""
}

# Limpieza y tiempo de borrado
$cleanupSeconds = $null
if ($CleanupRemote) {
  Write-Host "[INFO] Cleaning up remote run folder..."
  if (Test-Path -LiteralPath $remoteRun) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Remove-Item -LiteralPath $remoteRun -Recurse -Force
    $sw.Stop()
    $cleanupSeconds = [math]::Round($sw.Elapsed.TotalSeconds, 2)
    Append-TextLine -Path $ResultsPath -Text ("[CLEANUP] Remote folder deleted in {0:n2}s" -f $cleanupSeconds)
  } else {
    Append-TextLine -Path $ResultsPath -Text "[CLEANUP] Remote folder not found (nothing to delete)"
  }
} else {
  Append-TextLine -Path $ResultsPath -Text "[CLEANUP] Skipped (CleanupRemote = false)"
}

# Volcar filas CSV (con cleanup_sec en la última fila)
if ($ExportCsv -and $csvRows.Count -gt 0) {
  # Preparar cabecera si no existe
  $HeaderNew = "timestamp,run,files,total_MB,up_sec,up_mbps,down_sec,down_mbps,threads,cleanup_sec"
  if (-not (Test-Path -LiteralPath $CsvPath)) {
    $HeaderNew | Out-File -LiteralPath $CsvPath -Encoding UTF8 -Force
  } else {
    $first = (Get-Content -LiteralPath $CsvPath -TotalCount 1)
    if ($first -ne $HeaderNew) {
      # Re-migración defensiva
      $bak = "$CsvPath.bak_$timestamp"
      Rename-Item -LiteralPath $CsvPath -NewName $bak
      $HeaderNew | Out-File -LiteralPath $CsvPath -Encoding UTF8 -Force
      Append-TextLine -Path $ResultsPath -Text ("[MIGRATION] Header mismatch detected at write stage. Started new CSV. Old file saved as '{0}'" -f $bak)
    }
  }
  if ($cleanupSeconds -ne $null) { $csvRows[-1].cleanup_sec = [string]$cleanupSeconds }
  foreach ($r in $csvRows) {
    $csvLine = ('{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}' -f $r.timestamp,$r.run,$r.files,$r.total_MB,$r.up_sec,$r.up_mbps,$r.down_sec,$r.down_mbps,$r.threads,$r.cleanup_sec)
    Append-CsvLine -Path $CsvPath -CsvRow $csvLine
  }
}

Write-Host ""
Write-Host "=== Finished ==="
Write-Host ("TXT (per run): {0}" -f $ResultsPath)
if ($ExportCsv) { Write-Host ("CSV:           {0}" -f $CsvPath) }
