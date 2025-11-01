# Dossier contenant les archives .zip
$repertoire = "D:\Downloads\Archives"

# Parcours de tous les fichiers .zip du dossier
Get-ChildItem -Path $repertoire -Filter "*.zip" | ForEach-Object {
    $zipPath = $_.FullName
    $hashAttendu = $_.BaseName
    $hashCalcule = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash

    Write-Host "`n[$($_.Name)]"
    Write-Host "Hash (nom du fichier) : $hashAttendu"
    Write-Host "Hash calculé         : $hashCalcule"

    if ($hashCalcule -eq $hashAttendu) {
        Write-Host "→ Hash vérifié, extraction..."

        # Créer le dossier de destination personnalisé
        $destination = Join-Path -Path $repertoire -ChildPath $hashAttendu

        if (-not (Test-Path -Path $destination)) {
            New-Item -ItemType Directory -Path $destination | Out-Null
        }

        Expand-Archive -Path $zipPath -DestinationPath $destination -Force
        Write-Host "   Extraction réussie dans $destination"
    }
    else {
        Write-Host "→ Hash non valide, renommage..."

        # Nouveau nom avec extension .KO
        $nouveauNom = $hashAttendu + ".KO"
        $nouveauChemin = Join-Path -Path $repertoire -ChildPath $nouveauNom

        Rename-Item -Path $zipPath -NewName $nouveauNom
        Write-Host "   Fichier renommé en $nouveauChemin"
    }
}
