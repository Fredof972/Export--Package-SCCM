<#
.SYNOPSIS
  Exporte des applications MECM depuis une liste de noms et produit 1 ZIP final par application :
  <OutputPath>\<Nom>.zip contenant le ZIP paramètres (+) le dossier <Nom>_files générés par Export-CMApplication.

.PARAMETERS
  -SiteServer  : Nom FQDN/NetBIOS du serveur de site (SMS Provider)
  -SiteCode    : Code site (ex: P01)
  -OutputPath  : Répertoire de sortie des ZIP finaux
  -ListPath    : Fichier texte avec 1 nom d’application par ligne (lignes vides/commençant par # ignorées)
  -ExactMatch  : Si présent, correspondance exacte ; sinon LIKE (équivalent *Nom*)
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
$ProgressPreference = 'SilentlyContinue'  # évite la barre de progression lente

# ---------- 1) Charger la liste (FileSystem forcé) ----------
if (-not (Test-Path -LiteralPath $ListPath)) { throw "ListPath introuvable: $ListPath" }
$fsListPath = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $ListPath).Path

$wanted = Get-Content -LiteralPath $fsListPath |
          ForEach-Object { $_.Trim() } |
          Where-Object { $_ -and -not $_.StartsWith('#') } |
          Select-Object -Unique

if (-not $wanted) { throw "La liste est vide: $ListPath" }

# ---------- 2) Connexion au site MECM ----------
Import-Module ConfigurationManager -ErrorAction Stop
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
  New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}
Set-Location ("{0}:" -f $SiteCode)

# ---------- 3) Résoudre OutputPath côté FileSystem ----------
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$fsOutput = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $OutputPath).Path

# ---------- 4) Préparer la sélection d'apps ----------
# Récupérer toutes les apps une fois (rapide) puis filtrer en mémoire
$all = Get-CMApplication -Fast

if ($ExactMatch) {
  $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $wanted | ForEach-Object { [void]$set.Add($_) }
  $toExport = $all | Where-Object { $set.Contains($_.LocalizedDisplayName) }
} else {
  # LIKE multi-modèles via regex compilées (équivalent *mot*, insensible à la casse)
  $regexes = $wanted | ForEach-Object {
    New-Object System.Text.RegularExpressions.Regex([Regex]::Escape($_), 'IgnoreCase')
  }
  $toExport = $all | Where-Object {
    $name = $_.LocalizedDisplayName
    foreach ($rx in $regexes) { if ($rx.IsMatch($name)) { return $true } }
    return $false
  }
}

if (-not $toExport) {
  Write-Warning "Aucune application ne correspond à la liste."
  Set-Location C:\
  return
}

# ---------- Helpers ----------
function Sanitize([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
  $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
  ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

# 5) Export : 1 ZIP final par appli
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

function Sanitize([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
  $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
  ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

foreach ($app in $toExport) {
  $name     = $app.LocalizedDisplayName
  $safe     = Sanitize $name
  $workDir  = Join-Path $OutputPath $safe
  $paramZipName = "$safe.zip"
  $finalZip = Join-Path $OutputPath "$safe.zip"

  try {
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force }
    New-Item -ItemType Directory -Path $workDir | Out-Null

    # 1) Export natif dans le dossier de travail (OK sans FileSystem::)
    Export-CMApplication -InputObject $app -Path $workDir -FileName $paramZipName -Force -ErrorAction Stop

    # 2) Zipper tout le dossier de travail en un seul zip final
    if (Test-Path -LiteralPath $finalZip) { Remove-Item -LiteralPath $finalZip -Force }
    Compress-Archive -Path (Join-Path $workDir '*') -DestinationPath $finalZip -Force

    # 3) Nettoyage du dossier de travail
    Remove-Item -LiteralPath $workDir -Recurse -Force

    Write-Host "[OK] $name -> $finalZip" -ForegroundColor Green
  }
  catch {
    Write-Warning "[KO] $name : $($_.Exception.Message)"
    try { if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force } } catch {}
  }
}

Set-Location C:\
Write-Host "Terminé. Exports: $OutputPath" -ForegroundColor Cyan
# ---------- Fin ----------
Set-Location C:\
Write-Host "Terminé. Exports: $fsOutput" -ForegroundColor Cyan