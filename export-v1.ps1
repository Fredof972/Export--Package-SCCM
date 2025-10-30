<#
Exporte plusieurs applications MECM (ConfigMgr) depuis une liste.
Chaque application est exportée dans un dossier dédié : 
<OutputPath>\<LocalizedDisplayName>\

Exemple : 
D:\FF\Export\7-Zip\7-Zip.zip
D:\FF\Export\7-Zip\7-Zip_files\
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,
  [Parameter(Mandatory)] [string] $SiteCode,
  [Parameter(Mandatory)] [string] $OutputPath,   # ex: D:\FF\Export
  [Parameter(Mandatory)] [string] $ListPath,     # fichier texte : 1 nom par ligne
  [switch] $ExactMatch
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
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

 

# ---------- 5) Export par application : <OutputPath>\<LocalizedDisplayName>\ ----------
function Sanitize([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
    $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
    ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

foreach ($app in $toExport) {
    $name = $app.LocalizedDisplayName
    $safe = Sanitize $name

    # Dossier cible de cette application : <OutputPath>\<LocalizedDisplayName>\
    $appDir = Join-Path $OutputPath $safe

    # Si un dossier de même nom existe déjà (nom dupliqué), on suffixe avec le CI_ID pour éviter l’écrasement
    if (Test-Path -LiteralPath $appDir) {
        $appDir = Join-Path $OutputPath ("{0}__{1}" -f $safe, $app.CI_ID)
    }

    try {
        if (Test-Path -LiteralPath $appDir) { Remove-Item -LiteralPath $appDir -Recurse -Force }
        New-Item -ItemType Directory -Path $appDir | Out-Null

        # IMPORTANT : sortir du provider CMSite pour éviter le bug "Path specified more than once"
        Push-Location C:\

        # Export natif MECM dans ce dossier (produit "<Nom>.zip" + "<Nom>_files\")
        Export-CMApplication -InputObject $app -Path $appDir -FileName ("{0}.zip" -f $safe) -Force -ErrorAction Stop

        Pop-Location

        Write-Host "[OK] $name -> $appDir" -ForegroundColor Green
    }
    catch {
        Write-Warning "[KO] $name : $($_.Exception.Message)"
        try { if (Test-Path -LiteralPath $appDir) { Remove-Item -LiteralPath $appDir -Recurse -Force } } catch {}
    }
}

 

Set-Location C:\

Write-Host "Terminé. Exports: $OutputPath" -ForegroundColor Cyan
