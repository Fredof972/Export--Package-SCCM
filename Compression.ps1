
# Chemin du dossier à compresser (sans *)
$Path = "C:\Riot Games\Riot Client"

# Extraction du dossier parent
$PathZip = Split-Path -Path $Path -Parent

# Construction du nom de fichier zip basé sur le nom du dossier
$NomZip = (Split-Path -Path $Path -Leaf) + ".zip"
$Destination = Join-Path -Path $PathZip -ChildPath $NomZip

# Vérifie si l'archive existe déjà
If (-not (Test-Path $Destination)) {
    Compress-Archive -Path "$Path\*" -CompressionLevel Optimal -DestinationPath $Destination
    Write-Host "Archive créée : "+ $Destination
} else {
    Compress-Archive -Path "$Path\*" -CompressionLevel Optimal -DestinationPath $Destination -Force
    Write-Host "Déjà créée : "+ $Destination
}

# Calcul du hash SHA256
$hash = Get-FileHash $Destination -Algorithm SHA256
Write-Host "Hash SHA256 : " $hash.Hash
$NewName = "$($hash.Hash).zip"
Rename-Item -Path $Destination -NewName $NewName
Write-Host "Fichier renommé en : $NewName"