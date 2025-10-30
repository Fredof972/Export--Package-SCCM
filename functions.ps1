<# ====================== LOGGING (CMTrace-friendly) ====================== #>

# Initialise (crée le dossier si besoin). Retourne le chemin du log.
function New-Log {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Path
  )
  $dir = Split-Path -Parent $Path
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
  if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType File -Path $Path | Out-Null }
  return $Path
}

# Ecrit au format CMTrace : Date<TAB>Time<TAB>Component<TAB>Context<TAB>Type<TAB>Thread<TAB>Message
# Type: 1=INFO, 2=WARN, 3=ERROR
function Write-Log {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO',
    [Parameter(Mandatory)][string]$LogPath,
    [string]$Component = 'Script',
    [string]$Context = ''
  )
  $dt = Get-Date
  $type = switch ($Level) { 'INFO' {1} 'WARN' {2} 'ERROR' {3} }
  $line = ('{0}' -f (
    "{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f `
      $dt.ToString('MM-dd-yyyy'), `
      $dt.ToString('HH:mm:ss.fff'), `
      $Component, `
      $Context, `
      $type, `
      [Threading.Thread]::CurrentThread.ManagedThreadId, `
      $Message
  ))
  Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

<# ====================== HASH (SHA-256) ====================== #>

# Manifeste CSV: RelativePath,SizeBytes,SHA256
function New-HashManifest {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$RootPath,
    [string]$OutFile = (Join-Path $RootPath 'hash_manifest.csv'),
    [string]$Include = '*',         # filtre glob
    [switch]$Recurse
  )
  if (-not (Test-Path -LiteralPath $RootPath)) { throw "RootPath introuvable: $RootPath" }
  $files = Get-ChildItem -LiteralPath $RootPath -File -Filter $Include -Recurse:$Recurse
  "RelativePath,SizeBytes,SHA256" | Out-File -LiteralPath $OutFile -Encoding UTF8
  foreach ($f in $files) {
    $rel = $f.FullName.Substring($RootPath.TrimEnd('\','/').Length).TrimStart('\','/')
    $hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash.ToUpperInvariant()
    """$rel"",$($f.Length),$hash" | Out-File -LiteralPath $OutFile -Append -Encoding UTF8
  }
  return $OutFile
}

# Vérifie un manifeste CSV (colonnes: RelativePath,SizeBytes,SHA256)
# Retourne un objet avec : MissingFiles, ExtraFiles, HashMismatch
function Test-HashManifest {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$RootPath,
    [Parameter(Mandatory)][string]$ManifestCsv,
    [switch]$ThrowOnMismatch
  )
  if (-not (Test-Path -LiteralPath $RootPath)) { throw "RootPath introuvable: $RootPath" }
  if (-not (Test-Path -LiteralPath $ManifestCsv)) { throw "Manifest introuvable: $ManifestCsv" }

  $expected = Import-Csv -LiteralPath $ManifestCsv
  # carte attendue: relpath -> hash
  $map = @{}
  foreach ($e in $expected) { $map[$e.RelativePath] = $e }

  $missing     = New-Object System.Collections.ArrayList
  $mismatch    = New-Object System.Collections.ArrayList
  $seenRelpath = New-Object System.Collections.Generic.HashSet[string]

  foreach ($k in $map.Keys) {
    $dst = Join-Path $RootPath $k
    if (-not (Test-Path -LiteralPath $dst)) {
      [void]$missing.Add($k)
      continue
    }
    $h = (Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($h -ne $map[$k].SHA256.ToUpperInvariant()) {
      [void]$mismatch.Add([pscustomobject]@{ RelativePath=$k; Expected=$map[$k].SHA256; Actual=$h })
    }
    [void]$seenRelpath.Add($k)
  }

  # fichiers en plus (non listés dans le manifeste)
  $extra = Get-ChildItem -LiteralPath $RootPath -File -Recurse | ForEach-Object {
    $rel = $_.FullName.Substring($RootPath.TrimEnd('\','/').Length).TrimStart('\','/')
    if (-not $seenRelpath.Contains($rel) -and -not $map.ContainsKey($rel)) { $rel }
  }

  $result = [pscustomobject]@{
    MissingFiles = $missing
    HashMismatch = $mismatch
    ExtraFiles   = @($extra)
    IsOK         = ($missing.Count -eq 0 -and $mismatch.Count -eq 0)
  }

  if ($ThrowOnMismatch -and -not $result.IsOK) {
    throw "Hash check failed. Missing=$($missing.Count), Mismatch=$($mismatch.Count), Extra=$(@($extra).Count)"
  }
  return $result
}

<# ====================== ZIP (compress / extract) ====================== #>

# PS5/.NET 4.x : charger l’assembly FileSystem pour ZipFile
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null

function Compress-Directory {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$SourceDir,
    [Parameter(Mandatory)][string]$DestinationZip,
    [switch]$IncludeBaseDirectory,   # inclure le dossier racine lui-même
    [switch]$Force
  )
  if (-not (Test-Path -LiteralPath $SourceDir)) { throw "SourceDir introuvable: $SourceDir" }
  if (Test-Path -LiteralPath $DestinationZip) {
    if ($Force) { Remove-Item -LiteralPath $DestinationZip -Force } else { throw "Destination existe déjà: $DestinationZip (utilise -Force)" }
  }
  if ($IncludeBaseDirectory) {
    # CreateFromDirectory inclut le dossier base s’il est parent du contenu
    [IO.Compression.ZipFile]::CreateFromDirectory($SourceDir, $DestinationZip, [IO.Compression.CompressionLevel]::Optimal, $true)
  } else {
    # Zipper le contenu uniquement : ruser via dossier temporaire
    $tmp = New-Item -ItemType Directory -Path (Join-Path ([IO.Path]::GetTempPath()) ([IO.Path]::GetRandomFileName()))
    try {
      Get-ChildItem -LiteralPath $SourceDir -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $tmp.FullName -Recurse -Force
      }
      [IO.Compression.ZipFile]::CreateFromDirectory($tmp.FullName, $DestinationZip, [IO.Compression.CompressionLevel]::Optimal, $false)
    } finally {
      Remove-Item -LiteralPath $tmp.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
  return $DestinationZip
}

function Expand-Zip {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ZipPath,
    [Parameter(Mandatory)][string]$DestinationDir,
    [switch]$Force
  )
  if (-not (Test-Path -LiteralPath $ZipPath)) { throw "Zip introuvable: $ZipPath" }
  if (Test-Path -LiteralPath $DestinationDir) {
    if ($Force) { Remove-Item -LiteralPath $DestinationDir -Recurse -Force } else { throw "Destination existe déjà: $DestinationDir (utilise -Force)" }
  }
  New-Item -ItemType Directory -Path $DestinationDir | Out-Null
  [IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $DestinationDir)
  return $DestinationDir
}