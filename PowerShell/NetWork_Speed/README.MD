Done by: Nacho
PowerShell macro that checks Network speed.
You need to enter some values in the macro (edit it)

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


It creates a .CSV that IF EXISTS add the latest spped times to the CSV this way it contains a continuos CSV for multiple runs