# Export--Package-SCCM
# Applications modifiées depuis le 1er septembre 2025
.\Export-CMContent.ps1 -SiteServer "CM01.contoso.com" -SiteCode "P01" `
  -OutputRoot "D:\Exports\SCCM" -ObjectType Application -ModifiedAfter "2025-09-01"

# Deux packages par ID + ZIP par objet
.\Export-CMContent.ps1 -SiteServer "CM01.contoso.com" -SiteCode "P01" `
  -OutputRoot "\\filesrv\backup\SCCM" -ObjectType Package -Ids "P0100123","P0100ABC" -ZipPerObject

# Filtre par nom + DPs (best-effort) + parallélisme
.\Export-CMContent.ps1 -SiteServer "CM01" -SiteCode "P01" `
  -OutputRoot "D:\Exports" -ObjectType Both -NameLike "*7-Zip*","*Notepad++*" -IncludeDPs -MaxConcurrency 6
