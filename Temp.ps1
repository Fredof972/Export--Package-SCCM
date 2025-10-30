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