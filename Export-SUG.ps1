# Nom du Software Update Group à exporter
$SUGName = "NOM-DE-VOTRE-SUG"
# Chemin de sortie pour le CSV exporté
$CsvPath = "C:\Temp\Export-SUG.csv"

# Récupération de l'objet SUG
$SUG = Get-CMSoftwareUpdateGroup -Name $SUGName

# Récupération de la liste des updates contenues dans le SUG
$Updates = Get-CMSoftwareUpdate -UpdateGroupName $SUGName

# Sélection des propriétés les plus utiles à exporter
$Updates | Select-Object CI_ID, ArticleID, LocalizedDisplayName, DateCreated, Severity, IsExpired, IsSuperseded |
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nExport terminé : $CsvPath"
