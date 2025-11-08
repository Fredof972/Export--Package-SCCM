# Charger le module ConfigurationManager
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + 'ConfigurationManager.psd1')
Set-Location "<SiteCode>:"

# Paramètres
$sugName = "NomDeTaSUG"
$exportPath = "C:\SCCM-Export\$sugName"
New-Item -Path $exportPath -ItemType Directory -Force

# Extraction des updates de la SUG
$group = Get-CMSoftwareUpdateGroup -Name $sugName
$updates = Get-CMSoftwareUpdate -UpdateGroupId $group.Id
$updates | Select-Object CI_ID, ArticleID, LocalizedDisplayName, Date, ContentSourcePath |
  Export-Csv -Path "$exportPath\Updates.csv" -NoTypeInformation

# Copie des sources du package (optionnel)
$packageSource = "Dossier_source_du_package"
Copy-Item "$packageSource\*" $exportPath -Recurse
