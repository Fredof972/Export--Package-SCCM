<#
.SYNOPSIS
  Exporte des applications MECM depuis une liste de noms et produit 1 ZIP final par application :
  <OutputPath>\<Nom>.zip contenant le zip paramètres + le dossier <Nom>_files générés par Export-CMApplication.

.PARAMETERS
  -SiteServer  : FQDN/NetBIOS du SMS Provider
  -SiteCode    : Code site (ex: P01)
  -OutputPath  : Dossier de sortie
  -ListPath    : Fichier texte (1 nom par ligne, # = commentaire)
  -ExactMatch  : Si présent => correspondance exacte ; sinon LIKE (*Nom*)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,
  [Parameter(Mandatory)] [string] $SiteCode,
  [Parameter(Mandatory)] [string] $OutputPath,
  [Parameter(Mandatory)] [string] $ListPath,
  [switch] $ExactMatch
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------- 1) Lire la liste (FileSystem forcé) ----------
if (-not (Test-Path -LiteralPath $ListPath)) { throw "ListPath introuvable: $ListPath" }
$fsListPath = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $ListPath).Path

$wanted = Get-Content -LiteralPath $fsListPath |
          ForEach-Object { $_.Trim() } |
          Where-Object { $_ -and -not $_.StartsWith('#') } |
          Select-Object -Unique

if (-not $wanted) { throw "La liste est vide: $ListPath" }

# ---------- 2) Connexion MECM ----------
Import-Module ConfigurationManager -ErrorAction Stop
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
  New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}
Set-Location ("{0}:" -f $SiteCode)

# ---------- 3) Préparer l'Output ----------
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

# ---------- 4) Sélection des applications ----------
$all = Get-CMApplication -Fast

if ($ExactMatch) {
  $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $wanted | ForEach-Object { [void]$set.Add($_) }
  $toExport = $all | Where-Object { $set.Contains($_.LocalizedDisplayName) }
} else {
  $regexes = $wanted | ForEach-Object {
    New-Object System.Text.RegularExpressions.Regex([Regex]::Escape($_), 'IgnoreCase')
  }
  $toExport = $all | Where-Object {
    $n = $_.LocalizedDisplayName
    foreach ($rx in $regexes) { if ($rx.IsMatch($n)) { return $true } }
    return $false
  }
}

if (-not $toExport) {
  Write-Warning "Aucune application ne correspond à la liste."
  Set-Location C:\
  return
}

# ---------- Helpers ----------
function Sanitize([string]$s){
  if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
  $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
  ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

# ---------- 5) Export : dossier éponyme -> zip final -> cleanup ----------
foreach ($app in $toExport) {
  $name  = $app.LocalizedDisplayName
  $safe  = Sanitize $name

  # Dossier de travail : <OutputPath>\<Nom>\
  $workDir = Join-Path $OutputPath $safe

  # Nom du zip de paramètres produit par Export-CMApplication
  $paramZipName = "$safe.zip"

  # Zip final (gère collision par suffixe CI_ID si nécessaire)
  $finalZip = Join-Path $OutputPath "$safe.zip"
  if (Test-Path -LiteralPath $finalZip) {
    $finalZip = Join-Path $OutputPath ("{0}__{1}.zip" -f $safe, $app.CI_ID)
  }

  try {
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force }
    New-Item -ItemType Directory -Path $workDir | Out-Null

    # 1) Export natif dans le dossier de travail
    #    -> crée "<Nom>.zip" + "<Nom>_files\" DANS $workDir
    Export-CMApplication -InputObject $app -Path $workDir -FileName $paramZipName -Force -ErrorAction Stop

    # 2) Zipper tout le contenu du dossier de travail en un unique zip final "<Output>\<Nom>.zip"
    if (Test-Path -LiteralPath $finalZip) { Remove-Item -LiteralPath $finalZip -Force }
    Compress-Archive -Path (Join-Path $workDir '*') -DestinationPath $finalZip -Force

    # 3) Nettoyage : supprimer le dossier de travail
    Remove-Item -LiteralPath $workDir -Recurse -Force

    Write-Host "[OK] $name -> $finalZip" -ForegroundColor Green
  }
  catch {
    Write-Warning "[KO] $name : $($_.Exception.Message)"
    try { if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force } } catch {}
  }
}

# ---------- Fin ----------
Set-Location C:\
Write-Host "Terminé. Exports: $OutputPath" -ForegroundColor Cyan