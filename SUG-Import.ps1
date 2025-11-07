# À personnaliser :
# Chemin du CSV exporté (depuis l'autre SCCM)
$CsvPath = "C:\Temp\Export-SUG.csv"
# SUG cible (existant ou à créer)
$SUGName = "NOUVEAU-SUG"

# Charger les updates à importer
$Updates = Import-Csv -Path $CsvPath

# Vérifier si le SUG existe, sinon le créer
$SUG = Get-CMSoftwareUpdateGroup -Name $SUGName -ErrorAction SilentlyContinue
if (-not $SUG) {
    $SUG = New-CMSoftwareUpdateGroup -Name $SUGName
    Write-Host "Nouveau SUG créé : $SUGName"
} else {
    Write-Host "SUG existant trouvé : $SUGName"
}

# Collecte des updates à ajouter (par ArticleID / KB)
$UpdateCIs = foreach ($line in $Updates) {
    $upd = Get-CMSoftwareUpdate | Where-Object { $_.ArticleID -eq $line.ArticleID }
    if ($upd) { $upd.CI_ID }
    else { Write-Warning "KB non trouvé sur ce SUP: $($line.ArticleID)" }
}

# Ajout à la Software Update Group cible
if ($UpdateCIs.Count -gt 0) {
    Add-CMSoftwareUpdateToGroup -SoftwareUpdateId $UpdateCIs -UpdateGroupName $SUGName
    Write-Host "$($UpdateCIs.Count) updates ajoutées à $SUGName"
} else {
    Write-Warning "Aucune update à ajouter (aucune trouvée dans l'INFRA cible)."
}
