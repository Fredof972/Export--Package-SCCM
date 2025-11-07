# Nom de la SUG cible
$SUGName = "NOUVEAU-SUG"

# Chemin vers le CSV d'import
$CsvPath = "C:\chemin\vers\les\patchs.csv"

# Charger les updates à importer depuis le CSV
$Updates = Import-Csv -Path $CsvPath

# Vérifier si la SUG existe, sinon la créer
$SUG = Get-CMSoftwareUpdateGroup -Name $SUGName -ErrorAction SilentlyContinue
if (-not $SUG) {
    $SUG = New-CMSoftwareUpdateGroup -Name $SUGName
    Write-Host "Nouveau SUG créé : $SUGName"
} else {
    Write-Host "SUG existant trouvé : $SUGName"
}

# Collecte des updates à ajouter (par ArticleID)
$UpdateCIs = foreach ($line in $Updates) {
    $upd = Get-CMSoftwareUpdate -ArticleId $line.ArticleID
    if ($upd) { 
        $upd.CI_ID 
    } else { 
        Write-Warning "KB non trouvé dans SUP : $($line.ArticleID)" 
    }
}

# Ajouter les updates dans la Software Update Group cible
if ($UpdateCIs.Count -gt 0) {
    Add-CMSoftwareUpdateToGroup -SoftwareUpdateId $UpdateCIs -UpdateGroupName $SUGName
    Write-Host "$($UpdateCIs.Count) updates ajoutées à $SUGName"
} else {
    Write-Warning "Aucune update à ajouter (aucune trouvée dans le SUP)."
}
