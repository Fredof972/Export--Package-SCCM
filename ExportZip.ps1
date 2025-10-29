<#
.SYNOPSIS
  Exporte une liste d’applications SCCM (MECM) au format ZIP.
  1 application = 1 fichier ZIP.

.DESCRIPTION
  Lis un fichier contenant les noms d’applications,
  recherche chaque application dans SCCM, puis lance Export-CMApplication
  pour générer un ZIP par application dans le dossier spécifié.

.PARAMETERS
  -SiteServer  Nom du serveur de site (ex: CM01.contoso.com)
  -SiteCode    Code du site (ex: P01)
  -OutputPath  Dossier cible pour les fichiers ZIP
  -ListPath    Fichier texte contenant les noms d’applications
  -ExactMatch  Si présent : correspondance exacte, sinon recherche partielle (*nom*)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,
  [Parameter(Mandatory)] [string] $SiteCode,
  [Parameter(Mandatory)] [string] $OutputPath,
  [Parameter(Mandatory)] [string] $ListPath,
  [switch] $ExactMatch
)

$ErrorActionPreference  = 'Stop'
$ProgressPreference     = 'SilentlyContinue'  # accélère les boucles SCCM

# --- 1) Lecture de la liste d’applications ---
if (-not (Test-Path -LiteralPath $ListPath)) {
  throw "Fichier de liste introuvable : $ListPath"
}

# Forcer le provider FileSystem
$fsListPath = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $ListPath).Path
$wanted = Get-Content -LiteralPath $fsListPath |
          ForEach-Object { $_.Trim() } |
          Where-Object { $_ -and -not $_.StartsWith('#') } |
          Select-Object -Unique

if (-not $wanted) { throw "La liste est vide : $ListPath" }

# --- 2) Connexion au site ConfigMgr ---
Import-Module ConfigurationManager -ErrorAction Stop

if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
  New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}
Set-Location ("{0}:" -f $SiteCode)

# --- 3) Préparer le répertoire de sortie ---
if (-not (Test-Path -LiteralPath $OutputPath)) {
  New-Item -ItemType Directory -Path $OutputPath | Out-Null
}
$fsOutput = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $OutputPath).Path

# --- 4) Récupérer toutes les applications en mémoire ---
$allApps = Get-CMApplication -Fast

if ($ExactMatch) {
  $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $wanted | ForEach-Object { [void]$set.Add($_) }
  $toExport = $allApps | Where-Object { $set.Contains($_.LocalizedDisplayName) }
}
else {
  $regexes = $wanted | ForEach-Object {
    New-Object System.Text.RegularExpressions.Regex([Regex]::Escape($_), 'IgnoreCase')
  }
  $toExport = $allApps | Where-Object {
    $name = $_.LocalizedDisplayName
    foreach ($rx in $regexes) {
      if ($rx.IsMatch($name)) { return $true }
    }
    return $false
  }
}

if (-not $toExport) {
  Write-Warning "Aucune application trouvée correspondant à la liste."
  Set-Location C:\
  return
}

# --- 5) Export de chaque application ---
foreach ($app in $toExport) {
  $safe = ($app.LocalizedDisplayName -replace '[\\/:*?"<>|;]', '_')
  $zip  = "{0}.zip" -f $safe

  try {
    Export-CMApplication -InputObject $app -Path $fsOutput -FileName $zip -Force -ErrorAction Stop
    Write-Host "[OK] $($app.LocalizedDisplayName) exportée -> $zip" -ForegroundColor Green
  }
  catch {
    Write-Warning "[KO] $($app.LocalizedDisplayName) : $($_.Exception.Message)"
  }
}

# --- 6) Fin ---
Set-Location C:\
Write-Host ""
Write-Host "Terminé. Les exports sont dans : $OutputPath" -ForegroundColor Cyan
