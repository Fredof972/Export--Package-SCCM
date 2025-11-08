# Charger le module ConfigurationManager
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + 'ConfigurationManager.psd1')
Set-Location "<SiteCode>:"

# Paramètres
$sugName = "NomDeTaSUG"
$importPath = "C:\SCCM-Export\$sugName"

# Création de la SUG si absente
if (!(Get-CMSoftwareUpdateGroup -Name $sugName)) {
    New-CMSoftwareUpdateGroup -Name $sugName
}

# Import du CSV
$updates = Import-Csv -Path "$importPath\Updates.csv"

foreach ($update in $updates) {
    # Recherche le composant de mise à jour par ArticleID ou CI_ID
    $swUpdate = Get-CMSoftwareUpdate | Where-Object { $_.ArticleID -eq $update.ArticleID }
    if ($swUpdate) {
        Add-CMSoftwareUpdateToGroup -SoftwareUpdateGroupName $sugName -SoftwareUpdateId $swUpdate.CI_ID
    }
}
