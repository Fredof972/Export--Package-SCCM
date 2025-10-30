# $app est ton objet application, $appExportDir est le dossier "<OutputPath>\<LocalizedDisplayName>"
$zipName = ($app.LocalizedDisplayName -replace '[\\/:*?"<>|;]', '_') + ".zip"

# --- Neutraliser temporairement les PSDefaultParameterValues sur Export-CMApplication ---
$__backupDefaults = @{}
foreach ($k in @('Export-CMApplication:Path','Export-CMApplication:FileName','Export-CMApplication:InputObject','Export-CMApplication:Name')) {
    if ($PSDefaultParameterValues.ContainsKey($k)) {
        $__backupDefaults[$k] = $PSDefaultParameterValues[$k]
        $PSDefaultParameterValues.Remove($k) | Out-Null
    }
}

try {
    # Sort du provider CMSite (évite aussi les confits de provider)
    Push-Location C:\

    # IMPORTANT : -Path doit être un chemin FileSystem simple, pas qualifié, pas de double quotes exotiques
    Export-CMApplication -InputObject $app -Path $appExportDir -FileName $zipName -Force -ErrorAction Stop
}
finally {
    Pop-Location
    # Restaure les defaults s'il y en avait
    foreach ($k in $__backupDefaults.Keys) {
        $PSDefaultParameterValues[$k] = $__backupDefaults[$k]
    }
}


====================================================

# Prérequis: tu as déjà Set-Location "$SiteCode:" et $app = Get-CMApplication -Name "<Nom>"
$appExportDir = Join-Path "D:\_FF\Export" (($app.LocalizedDisplayName -replace '[\\/:*?"<>|;]', '_'))
if (Test-Path $appExportDir) { Remove-Item -Recurse -Force $appExportDir }
New-Item -ItemType Directory -Path $appExportDir | Out-Null

# Neutralise les defaults
$__backupDefaults = @{}
foreach ($k in @('Export-CMApplication:Path','Export-CMApplication:FileName','Export-CMApplication:InputObject','Export-CMApplication:Name')) {
    if ($PSDefaultParameterValues.ContainsKey($k)) {
        $__backupDefaults[$k] = $PSDefaultParameterValues[$k]
        $PSDefaultParameterValues.Remove($k) | Out-Null
    }
}

try {
    Push-Location C:\
    Export-CMApplication -InputObject $app -Path $appExportDir -FileName ( ($app.LocalizedDisplayName -replace '[\\/:*?"<>|;]', '_') + ".zip") -Force -ErrorAction Stop
}
finally {
    Pop-Location
    foreach ($k in $__backupDefaults.Keys) { $PSDefaultParameterValues[$k] = $__backupDefaults[$k] }
}

Write-Host "OK -> $appExportDir"

==========================
<#
.SYNOPSIS
  Exporte plusieurs applications MECM (ConfigMgr) depuis une liste.
  Chaque application est exportée dans un dossier dédié : <OutputPath>\<LocalizedDisplayName>\

.EXAMPLE
  .\Export-MECMApps.ps1 -SiteServer CM01 -SiteCode LAB -OutputPath D:\FF\Export -ListPath .\apps.txt
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,
  [Parameter(Mandatory)] [string] $SiteCode,
  [Parameter(Mandatory)] [string] $OutputPath,
  [Parameter(Mandatory)] [string] $ListPath,
  [switch] $ExactMatch,
  [string] $LogPath = "$($env:TEMP)\MECM_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-Log {
    param([string]$msg, [string]$level = 'INFO')
    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$time][$level] $msg"
    $line | Out-File -Append -Encoding UTF8 -FilePath $LogPath
    if ($level -eq 'ERROR') { Write-Warning $msg }
    elseif ($level -eq 'INFO') { Write-Host $msg -ForegroundColor Cyan }
    else { Write-Host $msg }
}

# 1) Charger la liste
if (-not (Test-Path -LiteralPath $ListPath)) { throw "ListPath introuvable: $ListPath" }
$wanted = Get-Content -LiteralPath $ListPath |
          ForEach-Object { $_.Trim() } |
          Where-Object { $_ -and -not $_.StartsWith('#') } |
          Select-Object -Unique
if (-not $wanted) { throw "La liste est vide: $ListPath" }

# 2) Connexion MECM
Import-Module ConfigurationManager -ErrorAction Stop
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}
Set-Location ("{0}:" -f $SiteCode)

# 3) Vérif du dossier Output
if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# 4) Récupération et filtrage
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
    Write-Log "Aucune application trouvée correspondant à la liste." "WARN"
    Set-Location C:\
    return
}

function Sanitize([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
    $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
    ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

# 5) Export des applications
foreach ($app in $toExport) {
    $name = $app.LocalizedDisplayName
    $safe = Sanitize $name
    $appExportDir = Join-Path $OutputPath $safe
    $zipName = "$safe.zip"
    $zipPath = Join-Path $appExportDir $zipName

    if (Test-Path -LiteralPath $appExportDir) { Remove-Item -LiteralPath $appExportDir -Recurse -Force }
    New-Item -ItemType Directory -Path $appExportDir | Out-Null

    try {
        Push-Location C:\
        Export-CMApplication -InputObject $app -Path $appExportDir -FileName $zipName -Force -ErrorAction Stop
        Pop-Location
        Write-Log "[OK] $name exporté vers $zipPath"
    }
    catch {
        Write-Log "[KO] $name : $($_.Exception.Message)" "ERROR"
        try { if (Test-Path -LiteralPath $appExportDir) { Remove-Item -LiteralPath $appExportDir -Recurse -Force } } catch {}
    }
}

Set-Location C:\
Write-Log "Terminé. Exports enregistrés dans : $OutputPath"
Write-Host "Journal : $LogPath" -ForegroundColor Gray
