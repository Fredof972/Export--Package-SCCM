function New-HashManifestXml {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$RootPath,     # Dossier racine à inventorier
    [string]$OutFile = (Join-Path $RootPath 'hash_manifest.xml'),
    [switch]$Recurse,                            # Inclure la descente récursive
    [string]$ApplicationName,                    # (optionnel) ajouté dans l'entête
    [string]$Comment                             # (optionnel) ajouté dans l'entête
  )

  if (-not (Test-Path -LiteralPath $RootPath)) { throw "RootPath introuvable: $RootPath" }

  $files = if ($Recurse) {
    Get-ChildItem -LiteralPath $RootPath -File -Recurse
  } else {
    Get-ChildItem -LiteralPath $RootPath -File
  }

  # Pré-calculs
  $totalBytes = ($files | Measure-Object -Property Length -Sum).Sum
  $utcNow = [DateTime]::UtcNow.ToString('o')

  # Construction XML
  $xml = New-Object System.Xml.XmlDocument
  $null = $xml.AppendChild($xml.CreateXmlDeclaration("1.0","utf-8",$null))
  $root = $xml.CreateElement("HashManifest")
  $null = $xml.AppendChild($root)

  # Attributs racine
  $root.SetAttribute("RootPath", (Resolve-Path -LiteralPath $RootPath).Path)
  $root.SetAttribute("GeneratedUtc", $utcNow)
  $root.SetAttribute("Algorithm", "SHA-256")
  $root.SetAttribute("FileCount", $files.Count)
  $root.SetAttribute("TotalBytes", [string]$totalBytes)

  if ($ApplicationName) { $root.SetAttribute("ApplicationName", $ApplicationName) }
  if ($Comment)         { $root.SetAttribute("Comment",         $Comment) }

  # Métadonnées host (facultatif mais utile)
  $meta = $xml.CreateElement("Environment")
  $meta.SetAttribute("Machine", $env:COMPUTERNAME)
  $meta.SetAttribute("User", $env:USERNAME)
  $meta.SetAttribute("PSVersion", $PSVersionTable.PSVersion.ToString())
  $null = $root.AppendChild($meta)

  # Liste des fichiers
  $filesNode = $xml.CreateElement("Files")
  $null = $root.AppendChild($filesNode)

  $rootNorm = $RootPath.TrimEnd('\','/')
  foreach ($f in $files) {
    $rel = $f.FullName.Substring($rootNorm.Length).TrimStart('\','/')
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $f.FullName).Hash.ToUpperInvariant()

    $fileNode = $xml.CreateElement("File")
    $fileNode.SetAttribute("RelativePath", $rel)
    $fileNode.SetAttribute("SizeBytes",    [string]$f.Length)
    $fileNode.SetAttribute("SHA256",       $hash)

    # Ajout d’un nœud <Time> avec dates utiles
    $timeNode = $xml.CreateElement("Time")
    $timeNode.SetAttribute("CreatedUtc",   ($f.CreationTimeUtc.ToString('o')))
    $timeNode.SetAttribute("ModifiedUtc",  ($f.LastWriteTimeUtc.ToString('o')))
    $timeNode.SetAttribute("AccessedUtc",  ($f.LastAccessTimeUtc.ToString('o')))
    $null = $fileNode.AppendChild($timeNode)

    $null = $filesNode.AppendChild($fileNode)
  }

  # Sauvegarde
  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Indent = $true
  $settings.IndentChars = "  "
  $settings.Encoding = New-Object System.Text.UTF8Encoding($false)  # UTF-8 sans BOM

  $writer = [System.Xml.XmlWriter]::Create($OutFile, $settings)
  try   { $xml.Save($writer) }
  finally { $writer.Dispose() }

  return $OutFile
}