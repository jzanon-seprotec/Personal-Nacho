param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$FilePath
)

Add-Type -AssemblyName PresentationFramework
function Show-Ok([string]$Text, [string]$Caption = "INFO", [string]$Icon = "Information") {
    [System.Windows.MessageBox]::Show($Text, $Caption, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::$Icon) | Out-Null
}

if (-not (Test-Path -LiteralPath $FilePath)) {
    Show-Ok "FILE NOT FOUND" "ERROR" "Error"
    exit 1
}

$dir  = [System.IO.Path]::GetDirectoryName($FilePath)
$base = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$ext  = ([System.IO.Path]::GetExtension($FilePath)).ToUpperInvariant()

if ($ext -ne ".PPF" -and $ext -ne ".TPF") {
    Show-Ok "FILE CAN'T BE PROCESSED, IT IS NOT A .PPF OR .TPF FILE" "ERROR" "Error"
    exit 0
}

function New-TempWorkdir([string]$hint) {
    $p = Join-Path ([System.IO.Path]::GetTempPath()) ("SEPROFIX_{0}_{1}" -f $hint, [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $p | Out-Null
    return $p
}

function Expand-ArchiveSafe($inputPath, $outputDir) {
    $tempZip = "$inputPath.zip"
    Copy-Item -LiteralPath $inputPath -Destination $tempZip -Force
    Expand-Archive -Path $tempZip -DestinationPath $outputDir -Force
    Remove-Item -LiteralPath $tempZip -Force
}

function Repack-Zip($sourceDir, $finalPath) {
    $tempZip = "$finalPath.zip"
    if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
    Compress-Archive -Path (Join-Path $sourceDir '*') -DestinationPath $tempZip
    if (Test-Path $finalPath) { Remove-Item $finalPath -Force }
    Rename-Item -Path $tempZip -NewName (Split-Path -Leaf $finalPath)
}

switch ($ext) {

    ".PPF" {
        $workPath = Join-Path $dir ("{0}_SEPROFIX{1}" -f $base, $ext)
        Copy-Item -LiteralPath $FilePath -Destination $workPath -Force

        $tmp = New-TempWorkdir $base
        Expand-ArchiveSafe -inputPath $workPath -outputDir $tmp

        $fr1 = Get-ChildItem -Path $tmp -Filter *.FR1 -File -ErrorAction SilentlyContinue
        $en1 = Get-ChildItem -Path $tmp -Filter *.EN1 -File -ErrorAction SilentlyContinue
        if (($fr1.Count + $en1.Count) -eq 0) {
            Show-Ok "IT IS NOT NECCESARY TO CHANGE ANYTHING ON THIS FILE; WILL WORK IN STUDIO AS IS"
            Remove-Item -LiteralPath $tmp -Recurse -Force
            exit 0
        }

        foreach ($f in $fr1) { Rename-Item $f.FullName ($f.FullName -replace '\.FR1$', '.FRA') }
        foreach ($f in $en1) { Rename-Item $f.FullName ($f.FullName -replace '\.EN1$', '.ENG') }

        $prjPath = Join-Path $tmp "$base.PRJ"
        if (-not (Test-Path $prjPath)) {
            $prj = Get-ChildItem $tmp -Filter *.PRJ -File | Select-Object -First 1
            if ($prj) { $prjPath = $prj.FullName }
        }

        if ($prjPath) {
            $content = Get-Content $prjPath -Raw
            $content = $content -replace 'SourceLanguage=31753', 'SourceLanguage=2057'
            $content = $content -replace 'SourceLanguage=31756', 'SourceLanguage=1036'
            # ✅ Añadir _SEPROFIX si no está al final
            $content = [regex]::Replace($content, '(?im)^ProjectName=(.+)', {
                param($m)
                $val = $m.Groups[1].Value.Trim()
                if ($val -notmatch '_SEPROFIX$') {
                    return "ProjectName=$val`_SEPROFIX"
                } else {
                    return $m.Value
                }
            })
            Set-Content -Path $prjPath -Value $content

            $newPrjName = "$base`_SEPROFIX.PRJ"
            if ((Split-Path -Leaf $prjPath) -ne $newPrjName) {
                Rename-Item $prjPath -NewName $newPrjName
            }
        }

        Repack-Zip -sourceDir $tmp -finalPath $workPath
        Remove-Item $tmp -Recurse -Force
        Show-Ok "PROCESS COMPLETED SUCCESSFULLY" "DONE"
    }

    ".TPF" {
        if ($base -notmatch '_SEPROFIX$') {
            Show-Ok "THIS IS NOT A SEPROFIX TPF FILE. NOTHING TO CHANGE"
            exit 0
        }

        $newBase = $base -replace '_SEPROFIX$', ''
        $workPath = Join-Path $dir ("{0}{1}" -f $newBase, $ext)
        Copy-Item $FilePath $workPath -Force

        $tmp = New-TempWorkdir $newBase
        Expand-ArchiveSafe -inputPath $workPath -outputDir $tmp

        $fra = Get-ChildItem $tmp -Filter *.FRA -File -ErrorAction SilentlyContinue
        $eng = Get-ChildItem $tmp -Filter *.ENG -File -ErrorAction SilentlyContinue
        foreach ($f in $fra) { Rename-Item $f.FullName ($f.FullName -replace '\.FRA$', '.FR1') }
        foreach ($f in $eng) { Rename-Item $f.FullName ($f.FullName -replace '\.ENG$', '.EN1') }

        $prj = Get-ChildItem $tmp -Filter *.PRJ -File | Select-Object -First 1
        if ($prj) {
            $prjPath = $prj.FullName
            $content = Get-Content $prjPath -Raw
            $content = $content -replace 'SourceLanguage=2057', 'SourceLanguage=31753'
            $content = $content -replace 'SourceLanguage=1036', 'SourceLanguage=31756'
            # ✅ Quitar _SEPROFIX solo si está al final
            $content = [regex]::Replace($content, '(?im)^ProjectName=(.+)', {
                param($m)
                $val = $m.Groups[1].Value.Trim() -replace '_SEPROFIX$', ''
                return "ProjectName=$val"
            })
            Set-Content -Path $prjPath -Value $content

            if ($prj.Name -match '_SEPROFIX\.PRJ$') {
                Rename-Item $prjPath -NewName ($prj.Name -replace '_SEPROFIX\.PRJ$', '.PRJ')
            }
        }

        Repack-Zip -sourceDir $tmp -finalPath $workPath
        Remove-Item $tmp -Recurse -Force
        Show-Ok "PROCESS COMPLETED SUCCESSFULLY" "DONE"
    }
}
