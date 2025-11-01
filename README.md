# Export--Package-SCCM
.\apps.txt  (exemple)
7-Zip
Google Chrome
Notepad++

.\Export-CMApps-FromList.ps1 `
  -SiteServer "CM01.contoso.com" -SiteCode "P01" `
  -OutputPath "D:\RepoExports" `
  -ListPath ".\apps.txt" `
  -ExactMatch   # retire ce switch si tu veux un LIKE *nom*

