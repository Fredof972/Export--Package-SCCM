# Nom de la SUG à exporter
$SUGName = "NOUVEAU-SUG"

# Chemin de destination du fichier CSV d'export
$ExportPath = "C:\chemin\vers\export-sug.csv"

# Récupérer la SUG
$SUG = Get-CMSoftwareUpdateGroup -Name $SUGName
if (-not $SUG) {
    Write-Warning "SUG '$SUGName' introuvable."
    exit
}

# Récupérer les mises à jour membres de la SUG
$Updates = Get-CMSoftwareUpdate -UpdateGroupName $SUGName

# Construire un objet avec les infos à exporter (ici ArticleID)
$ExportData = $Updates | Select-Object -Property ArticleID, Title, CI_ID, DateReleased

# Exporter en CSV
$ExportData | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "Export de la SUG '$SUGName' terminé dans : $ExportPath"
