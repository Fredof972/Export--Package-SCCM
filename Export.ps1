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

# ---------- 1) Charger la liste ----------
if (-not (Test-Path -LiteralPath $ListPath)) { throw "ListPath introuvable: $ListPath" }
$wanted = Get-Content -LiteralPath $ListPath |
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

# ---------- 3) Vérif Output ----------
if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# ---------- 4) Filtrage ----------
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

# ---------- Helper ----------
function Sanitize([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
    $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
    ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

# ---------- 5) Export par application ----------
foreach ($app in $toExport) {
    $name = $app.LocalizedDisplayName
    $safe = Sanitize $name

    # Créer un dossier dédié pour cette app
    $appExportDir = Join-Path $OutputPath $safe
    if (Test-Path -LiteralPath $appExportDir) { Remove-Item -LiteralPath $appExportDir -Recurse -Force }
    New-Item -ItemType Directory -Path $appExportDir | Out-Null

    $zipName = "$safe.zip"

    try {
        # ⚠️ Sort temporairement du provider CMSite pour éviter le conflit de -Path
        Push-Location C:\

        Export-CMApplication -InputObject $app -Path $appExportDir -FileName $zipName -Force -ErrorAction Stop

        Pop-Location

        Write-Host "[OK] $name exporté dans $appExportDir" -ForegroundColor Green
    }
    catch {
        Write-Warning "[KO] $name : $($_.Exception.Message)"
        try { if (Test-Path -LiteralPath $appExportDir) { Remove-Item -LiteralPath $appExportDir -Recurse -Force } } catch {}
    }
}

Set-Location C:\
Write-Host "Terminé. Exports enregistrés dans : $OutputPath" -ForegroundColor Cyan