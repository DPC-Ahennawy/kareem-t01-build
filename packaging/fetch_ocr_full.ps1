# Kareem T-01 - fetch REAL Tesseract engine + DLLs + ara/eng/fra traineddata
# Route B: downloads everything into  <app>\ocr\tesseract\  BEFORE packaging.
# Requires internet. Run once on the Windows build machine.
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$App     = Split-Path -Parent $PSScriptRoot          # ...\Kareem T-01
$OcrDir  = Join-Path $App "ocr\tesseract"
$TData   = Join-Path $OcrDir "tessdata"
New-Item -ItemType Directory -Force -Path $TData | Out-Null

Write-Host "== Kareem T-01 OCR fetch ==" -ForegroundColor Cyan
Write-Host "Target: $OcrDir"

# ---------------------------------------------------------------
# 1) Tesseract engine (portable). We use the UB-Mannheim installer
#    and extract the program files with 7-Zip (no admin install).
# ---------------------------------------------------------------
$exe = Join-Path $OcrDir "tesseract.exe"
if (Test-Path $exe) {
    Write-Host "tesseract.exe already present - skipping engine download." -ForegroundColor Green
} else {
    $tmp = Join-Path $env:TEMP "kt01_tess_setup.exe"
    # Pinned UB-Mannheim 5.3.x x64 installer (NSIS). Update the URL if the wiki changes.
    $url = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    Write-Host "Downloading Tesseract engine installer..."
    Invoke-WebRequest -Uri $url -OutFile $tmp

    # Need 7-Zip to extract the NSIS installer without running it.
    $sevenz = @("$env:ProgramFiles\7-Zip\7z.exe","${env:ProgramFiles(x86)}\7-Zip\7z.exe") |
              Where-Object { Test-Path $_ } | Select-Object -First 1
    $extract = Join-Path $env:TEMP "kt01_tess_x"
    if (Test-Path $extract) { Remove-Item $extract -Recurse -Force }

    if ($sevenz) {
        Write-Host "Extracting engine with 7-Zip..."
        & $sevenz x $tmp "-o$extract" -y | Out-Null
    } else {
        Write-Host "7-Zip not found - running the installer in silent mode to a temp dir..."
        # NSIS silent install into a temp folder, then copy files out.
        Start-Process $tmp -ArgumentList "/S","/D=$extract" -Wait
    }

    # Copy tesseract.exe + all DLLs (+ leptonica etc.) to our bundled folder.
    $src = Get-ChildItem $extract -Recurse -Filter "tesseract.exe" | Select-Object -First 1
    if (-not $src) { throw "tesseract.exe not found after extraction. Install 7-Zip and retry." }
    $srcDir = $src.Directory.FullName
    Write-Host "Copying engine + DLLs from $srcDir"
    Copy-Item (Join-Path $srcDir "tesseract.exe") $OcrDir -Force
    Get-ChildItem $srcDir -Filter *.dll | ForEach-Object { Copy-Item $_.FullName $OcrDir -Force }
    # some builds keep DLLs one level up
    Get-ChildItem $extract -Recurse -Filter *.dll | ForEach-Object {
        $dest = Join-Path $OcrDir $_.Name
        if (-not (Test-Path $dest)) { Copy-Item $_.FullName $dest -Force }
    }
}

# ---------------------------------------------------------------
# 2) Language data (official tessdata_best): eng, ara, fra
# ---------------------------------------------------------------
$base = "https://github.com/tesseract-ocr/tessdata_best/raw/main/"
foreach ($l in "eng","ara","fra") {
    $out = Join-Path $TData "$l.traineddata"
    if (Test-Path $out) { Write-Host "$l.traineddata present - skip"; continue }
    Write-Host "Downloading $l.traineddata ..."
    Invoke-WebRequest -Uri ($base + "$l.traineddata") -OutFile $out
}

# ---------------------------------------------------------------
# 3) Verify
# ---------------------------------------------------------------
Write-Host ""
$ok = $true
foreach ($f in @("tesseract.exe","tessdata\eng.traineddata","tessdata\ara.traineddata","tessdata\fra.traineddata")) {
    $p = Join-Path $OcrDir $f
    if (Test-Path $p) { Write-Host "  [OK] $f" -ForegroundColor Green }
    else { Write-Host "  [MISSING] $f" -ForegroundColor Red; $ok = $false }
}
$dlls = (Get-ChildItem $OcrDir -Filter *.dll).Count
Write-Host "  DLLs copied: $dlls"
if (-not $ok) { throw "OCR fetch incomplete - see [MISSING] above." }
Write-Host "OCR bundle ready." -ForegroundColor Cyan
