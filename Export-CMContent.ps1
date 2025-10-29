
[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,
  [Parameter(Mandatory)] [string] $SiteCode,
  [Parameter(Mandatory)] [string] $OutputPath,
  [Parameter(Mandatory)] [string] $ListPath,
  [switch] $ExactMatch     # sans ce switch: correspondances partielles (LIKE)
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'  # évite le rendu lent de la barre de progression

# 1) Lire la liste AVANT d'aller sur le drive du site (FileSystem forcé)
if (-not (Test-Path -LiteralPath $ListPath)) { throw "ListPath introuvable: $ListPath" }
$fsListPath = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $ListPath).Path
$wanted = Get-Content -LiteralPath $fsListPath |
          ForEach-Object { $_.Trim() } |
          Where-Object { $_ -and -not $_.StartsWith('#') } |
          Select-Object -Unique
if (-not $wanted) { throw "La liste est vide: $ListPath" }

# 2) Connexion CM
Import-Module ConfigurationManager -ErrorAction Stop
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
  New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}
Set-Location ("{0}:" -f $SiteCode)

# 3) Préparer sortie
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

# 4) Récupérer TOUTES les applis UNE FOIS (rapide) puis filtrer en mémoire
$all = Get-CMApplication -Fast

if ($ExactMatch) {
  # table de hachage pour lookup O(1)
  $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $wanted | ForEach-Object { [void]$set.Add($_) }
  $toExport = $all | Where-Object { $set.Contains($_.LocalizedDisplayName) }
}
else {
  # LIKE (multi-modèles) – on compile des regex une fois
  $regexes = $wanted | ForEach-Object {
    # équivaut à *mot* insensible à la casse
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

# 5) Export en chaîne (simple, stable)
foreach ($app in $toExport) {
  $safe = ($app.LocalizedDisplayName -replace '[\\/:*?"<>|;]', '_')
  $zip  = "{0}__{1}.zip" -f $safe, $app.CI_ID
  try {
    $app | Export-CMApplication -Path $OutputPath -FileName $zip -Force
    Write-Host "[OK] $($app.LocalizedDisplayName)" -ForegroundColor Green
  } catch {
    Write-Warning "[KO] $($app.LocalizedDisplayName) : $($_.Exception.Message)"
  }
}

Set-Location C:\
Write-Host "Terminé. Exports: $OutputPath" -ForegroundColor Cyan