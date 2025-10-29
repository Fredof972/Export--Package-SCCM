[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,   # ex: CM01.contoso.com
  [Parameter(Mandatory)] [string] $SiteCode,     # ex: P01
  [Parameter(Mandatory)] [string] $OutputPath,   # ex: D:\RepoExports
  [Parameter(Mandatory)] [string] $ListPath,     # fichier texte: 1 nom d'application par ligne
  [switch] $ExactMatch                          # sinon LIKE *nom*
)

$ErrorActionPreference = 'Stop'

# Prépa
Import-Module ConfigurationManager -ErrorAction Stop
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
  New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}
Set-Location "$SiteCode:`"
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

# Charge la liste (ignore lignes vides/commentaires)
$appNames = Get-Content -LiteralPath $ListPath | ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -Unique
if (-not $appNames) { throw "La liste est vide: $ListPath" }

function Sanitize([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
  $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
  ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

foreach ($name in $appNames) {
  try {
    # Résolution de l’objet application
    $apps = if ($ExactMatch) {
      Get-CMApplication | Where-Object { $_.LocalizedDisplayName -eq $name }
    } else {
      Get-CMApplication | Where-Object { $_.LocalizedDisplayName -like "*$name*" }
    }

    if (-not $apps) {
      Write-Warning "Aucune application trouvée pour: $name"
      continue
    }

    foreach ($app in $apps) {
      $safe = Sanitize $app.LocalizedDisplayName
      $zip  = "{0}__{1}.zip" -f $safe, $app.CI_ID
      Write-Host "Export: $($app.LocalizedDisplayName) -> $OutputPath\$zip"
      # Base demandée :
      $app | Export-CMApplication -Path $OutputPath -FileName $zip -Force
    }
  }
  catch {
    Write-Warning "Échec export [$name] : $($_.Exception.Message)"
  }
}

# Restaure l’emplacement
Set-Location C:\
Write-Host "Terminé. Exports dans: $OutputPath" -ForegroundColor Cyan
