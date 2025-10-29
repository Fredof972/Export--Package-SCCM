# Sécure: forcer un chemin FileSystem pour l’output
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$fsOutput = "Microsoft.PowerShell.Core\FileSystem::" + (Resolve-Path -LiteralPath $OutputPath).Path

function Sanitize([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "_empty" }
  $bad = [IO.Path]::GetInvalidFileNameChars() + '\','/',';',':','*','?','"','<','>','|'
  ($s.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
}

foreach ($app in $toExport) {
  $name    = $app.LocalizedDisplayName
  $safe    = Sanitize $name
  $workDir = Join-Path $fsOutput $safe           # OutputPath\<Nom>
  $tmpZip  = Join-Path $workDir  "$safe.export.zip"  # zip temporaire créé par Export-CMApplication
  $finalZip= Join-Path $fsOutput "$safe.zip"         # OutputPath\<Nom>.zip

  try {
    # 1) Créer le dossier de travail OutputPath\<Nom>\
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force }
    New-Item -ItemType Directory -Path $workDir | Out-Null

    # 2) Export-CMApplication DANS CE DOSSIER (zip temporaire), puis dézipper dans le même dossier
    Export-CMApplication -InputObject $app -Path $workDir -FileName (Split-Path -Leaf $tmpZip) -Force -ErrorAction Stop
    Expand-Archive -Path $tmpZip -DestinationPath $workDir -Force
    Remove-Item -LiteralPath $tmpZip -Force

    # 3) Zipper le dossier de travail en OutputPath\<Nom>.zip
    if (Test-Path -LiteralPath $finalZip) { Remove-Item -LiteralPath $finalZip -Force }
    Compress-Archive -Path (Join-Path $workDir '*') -DestinationPath $finalZip -Force

    # 4) Supprimer le dossier de travail
    Remove-Item -LiteralPath $workDir -Recurse -Force

    Write-Host "[OK] $name -> $finalZip" -ForegroundColor Green
  }
  catch {
    Write-Warning "[KO] $name : $($_.Exception.Message)"
    try { if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force } } catch {}
  }
}
