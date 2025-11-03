function Test-HashManifestXml {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$RootPath,     # Dossier à vérifier
    [Parameter(Mandatory)][string]$ManifestXml,  # Fichier XML généré par New-HashManifestXml
    [switch]$ThrowOnMismatch
  )

  if (-not (Test-Path -LiteralPath $RootPath))  { throw "RootPath introuvable: $RootPath" }
  if (-not (Test-Path -LiteralPath $ManifestXml)) { throw "Manifest introuvable: $ManifestXml" }

  [xml]$xml = Get-Content -LiteralPath $ManifestXml -Raw

  $missing  = New-Object System.Collections.ArrayList
  $mismatch = New-Object System.Collections.ArrayList
  $seen     = New-Object System.Collections.Generic.HashSet[string]

  $rootNorm = $RootPath.TrimEnd('\','/')
  $algo = $xml.HashManifest.Algorithm
  if ($algo -and $algo -ne 'SHA-256') {
    Write-Warning "Algorithme inattendu dans le manifeste: '$algo' (attendu: SHA-256)"
  }

  foreach ($fileNode in $xml.HashManifest.Files.File) {
    $rel = [string]$fileNode.RelativePath
    $exp = [string]$fileNode.SHA256
    $dst = Join-Path $RootPath $rel

    if (-not (Test-Path -LiteralPath $dst)) {
      [void]$missing.Add($rel)
      continue
    }

    $act = (Get-FileHash -Algorithm SHA256 -LiteralPath $dst).Hash.ToUpperInvariant()
    if ($act -ne $exp.ToUpperInvariant()) {
      [void]$mismatch.Add([pscustomobject]@{
        RelativePath = $rel
        Expected     = $exp
        Actual       = $act
      })
    }
    [void]$seen.Add($rel)
  }

  # fichiers supplémentaires non listés
  $extra = Get-ChildItem -LiteralPath $RootPath -File -Recurse | ForEach-Object {
    $rel = $_.FullName.Substring($rootNorm.Length).TrimStart('\','/')
    if (-not $seen.Contains($rel)) { $rel }
  }

  $result = [pscustomobject]@{
    IsOK         = ($missing.Count -eq 0 -and $mismatch.Count -eq 0)
    MissingFiles = $missing
    HashMismatch = $mismatch
    ExtraFiles   = @($extra)
    Algorithm    = ($xml.HashManifest.Algorithm)
    GeneratedUtc = ($xml.HashManifest.GeneratedUtc)
    FileCount    = [int]$xml.HashManifest.FileCount
  }

  if ($ThrowOnMismatch -and -not $result.IsOK) {
    throw "Hash check failed. Missing=$($missing.Count), Mismatch=$($mismatch.Count), Extra=$(@($extra).Count)"
  }

  return $result
}