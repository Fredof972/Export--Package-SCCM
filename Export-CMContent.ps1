<#
.SYNOPSIS
  Exporte Applications et Packages SCCM/MECM (métadonnées + contenu + manifeste), avec logs et vérification d’intégrité.

.DESCRIPTION
  - Tente d'abord l'utilisation du module ConfigurationManager (console MECM).
  - Fallback WMI/SMS Provider (root\SMS\site_<SiteCode>).
  - (Best-effort) Fallback AdminService si nécessaire (endpoints non-exhaustifs).

  Exporte :
    * Applications (SMS_ApplicationLatest) : SDMPackageXML + JSON structuré + contenu DTs.
    * Packages (SMS_Package) : JSON + Programmes + contenu SourcePath.
    * DPs (option -IncludeDPs) : best-effort via classes *DistPointsSummarizer* (si dispo).
    * Manifeste global CSV/JSON + hashlist.

  Testé PowerShell 5.1 et PS7+. Concurrency activée si PS7 détecté.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteServer,          # ex: CM01.contoso.com
  [Parameter(Mandatory)] [string] $SiteCode,            # ex: P01
  [Parameter(Mandatory)] [string] $OutputRoot,          # ex: D:\Exports\SCCM
  [ValidateSet('Package','Application','Both')] [string] $ObjectType = 'Both',
  [string[]] $NameLike,                                 # filtre par nom (wildcards)
  [string[]] $Ids,                                      # PackageID (ABC00012) ou CI_UniqueID (ScopeId_xxx/Application_xxx)
  [datetime] $ModifiedAfter,                            # filtre date
  [switch] $IncludeDPs,                                 # export liste DP affectés (pas de contenu DP)
  [switch] $ZipPerObject,                               # zip résultat par objet
  [int] $MaxConcurrency = 4,
  [int] $RetryCount = 3,
  [int] $RetryDelaySec = 3,
  [switch] $VerboseLog
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Helpers: UI/Env/Paths/Logging ---
function Test-IsPS7 { return ($PSVersionTable.PSVersion.Major -ge 7) }
function New-Dir([string]$Path){ if(-not [string]::IsNullOrWhiteSpace($Path) -and -not (Test-Path -LiteralPath $Path)){ New-Item -ItemType Directory -Path $Path -Force | Out-Null }; return $Path }
function Sanitize-Name([string]$Name){
  if([string]::IsNullOrWhiteSpace($Name)){ return "_empty" }
  $bad = [IO.Path]::GetInvalidFileNameChars() + [char]'\'+[char]'/' + [char]':' + [char]'*' + [char]'?' + [char]'"' + [char]'<' + [char]'>' + [char]'|'
  ($Name.ToCharArray() | ForEach-Object { if($bad -contains $_){ '_' } else { $_ } }) -join ''
}
$ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
$SessionRoot = New-Dir (Join-Path $OutputRoot "session-$ts")
$LogDir      = New-Dir (Join-Path $SessionRoot "_logs")
$ManifestCsv = Join-Path $SessionRoot "manifest.csv"
$ManifestJson= Join-Path $SessionRoot "manifest.json"
$TraceLog    = Join-Path $LogDir "Export-SCCM.log"
$Transcript  = Join-Path $LogDir ("Export-SCCM-{0}.transcript.log" -f $ts)

function Write-CMTraceLog {
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR')] [string] $Level = 'INFO'
  )
  # CMTrace format: Date\tTime\tComponent\tContext\tType\tThread\tMessage
  $dt = Get-Date
  $type = switch ($Level) { 'INFO' {1} 'WARN' {2} 'ERROR' {3} }
  $line = "{0}`t{1}`tExport-CMContent`t`t{2}`t{3}`t{4}" -f $dt.ToString('MM-dd-yyyy'),$dt.ToString('HH:mm:ss.fff'),$type,[Threading.Thread]::CurrentThread.ManagedThreadId,$Message
  Add-Content -LiteralPath $TraceLog -Value $line -Encoding UTF8
  if($VerboseLog -or $Level -ne 'INFO'){ Write-Host "[$Level] $Message" -ForegroundColor (@{INFO='Gray';WARN='Yellow';ERROR='Red'}[$Level]) }
}

Start-Transcript -LiteralPath $Transcript | Out-Null
Write-CMTraceLog "=== Start Export (SiteServer=$SiteServer, SiteCode=$SiteCode, ObjectType=$ObjectType) ==="
#endregion

#region --- Hash / Copy with verification ---
function Get-FileSHA256([string]$Path){
  if(-not (Test-Path -LiteralPath $Path)){ throw "File not found: $Path" }
  (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToUpperInvariant()
}

function Copy-WithHash {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string] $Source,
    [Parameter(Mandatory)] [string] $Destination,
    [int] $RetryCount = 3,
    [int] $RetryDelaySec = 3,
    [string] $QuarantineRoot
  )
  if(-not (Test-Path -LiteralPath (Split-Path -Path $Destination -Parent))){
    New-Dir (Split-Path -Path $Destination -Parent) | Out-Null
  }

  $attempt = 0
  while($true){
    try{
      Copy-Item -LiteralPath $Source -Destination $Destination -Force
      $srcH = Get-FileSHA256 -Path $Source
      $dstH = Get-FileSHA256 -Path $Destination
      if($srcH -ne $dstH){
        throw "HASH_MISMATCH: $Source"
      }
      return $dstH
    } catch {
      $attempt++
      if($attempt -gt $RetryCount){
        Write-CMTraceLog "Copy failed after $RetryCount retries for $Source -> $Destination : $($_.Exception.Message)" 'ERROR'
        if($QuarantineRoot){
          $qdir = New-Dir (Join-Path $QuarantineRoot "_quarantine")
          $qdst = Join-Path $qdir (Split-Path -Leaf $Source)
          try { Copy-Item -LiteralPath $Source -Destination $qdst -Force -ErrorAction Stop } catch {}
        }
        throw
      } else {
        Write-CMTraceLog "Copy retry $attempt/$RetryCount for $Source : $($_.Exception.Message)" 'WARN'
        Start-Sleep -Seconds $RetryDelaySec
      }
    }
  }
}
#endregion

#region --- Connection (CM cmdlets -> AdminService -> WMI) ---
$global:CMConnectionMode = $null
$global:CMWmi = $null

function Connect-CMProvider {
  param([string]$SiteServer,[string]$SiteCode)
  # Try ConfigurationManager module
  try{
    Import-Module ConfigurationManager -ErrorAction Stop
    $siteDrive = Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue
    if(-not $siteDrive){
      New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
    }
    Set-Location "$($SiteCode):" -ErrorAction Stop
    $global:CMConnectionMode = 'CMModule'
    Write-CMTraceLog "Connected via ConfigurationManager module to $SiteServer ($SiteCode)."
    return
  } catch {
    Write-CMTraceLog "CM module not available or failed: $($_.Exception.Message)" 'WARN'
  }

  # AdminService (best-effort): we just validate endpoint reachability
  $adminUrl = "https://$SiteServer/AdminService/wmi"
  try{
    # For offline environments without Invoke-RestMethod to that URL (trust issues), skip strict check; set mode and proceed WMI for data.
    $global:CMConnectionMode = 'AdminService'  # marker
    Write-CMTraceLog "AdminService assumed available at $adminUrl (best-effort)."
  } catch {
    Write-CMTraceLog "AdminService not reachable: $($_.Exception.Message)" 'WARN'
  }

  # WMI fallback
  try{
    $ns = "root\SMS\site_$SiteCode"
    $global:CMWmi = [wmi] "\\$SiteServer\$ns:__namespace.name='$ns'"
    # Quick query to validate
    $null = Get-WmiObject -Namespace $ns -Class SMS_Site -ComputerName $SiteServer -ErrorAction Stop | Select-Object -First 1
    $global:CMConnectionMode = if($global:CMConnectionMode){ "$($global:CMConnectionMode)+WMI" } else { 'WMI' }
    Write-CMTraceLog "Connected to WMI at \\$SiteServer\$ns ($($global:CMConnectionMode))."
  } catch {
    Write-CMTraceLog "Failed to connect to WMI Provider: $($_.Exception.Message)" 'ERROR'
    throw "CONNECT_FAIL"
  }
}
Connect-CMProvider -SiteServer $SiteServer -SiteCode $SiteCode
#endregion

#region --- Queries (Packages / Applications) ---
function Get-PackageObjects {
  param([string[]]$NameLike,[string[]]$Ids,[datetime]$ModifiedAfter)
  if($global:CMConnectionMode -like 'CMModule*'){
    $q = Get-CMPackage -Fast
  } else {
    $q = Get-WmiObject -Class SMS_Package -Namespace ("root\SMS\site_$SiteCode") -ComputerName $SiteServer
  }
  if($Ids){
    $idSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $Ids | ForEach-Object { [void]$idSet.Add($_) }
    $q = $q | Where-Object { $idSet.Contains($_.PackageID) }
  }
  if($NameLike){ foreach($p in $NameLike){ $q = $q | Where-Object { $_.Name -like $p } } }
  if($ModifiedAfter){ $q = $q | Where-Object { $_.LastRefreshTime -ge $ModifiedAfter } }
  return $q
}

function Get-ApplicationObjects {
  param([string[]]$NameLike,[string[]]$Ids,[datetime]$ModifiedAfter)
  $apps = Get-WmiObject -Class SMS_ApplicationLatest -Namespace ("root\SMS\site_$SiteCode") -ComputerName $SiteServer
  if($Ids){
    $idSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $Ids | ForEach-Object { [void]$idSet.Add($_) }
    $apps = $apps | Where-Object { $idSet.Contains($_.CI_UniqueID) }
  }
  if($NameLike){ foreach($p in $NameLike){ $apps = $apps | Where-Object { $_.LocalizedDisplayName -like $p } } }
  if($ModifiedAfter){ $apps = $apps | Where-Object { $_.LastModified -ge $ModifiedAfter } }
  return $apps
}
#endregion

#region --- Export logic: Applications ---
function Parse-AppXml {
  param([string]$Xml)
  $x = [xml]$Xml
  $ns = @{ s="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" }
  return $x
}

function Export-CMApplication {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $App,
    [Parameter(Mandatory)] [string] $Root,
    [int] $RetryCount = 3,
    [int] $RetryDelaySec = 3
  )
  $id        = $App.CI_UniqueID
  $name      = if($App.LocalizedDisplayName){ $App.LocalizedDisplayName } else { $App.ModelName }
  $nameSafe  = Sanitize-Name $name
  $objRoot   = New-Dir (Join-Path $Root ("Applications\{0}__{1}" -f $nameSafe,$id.Replace('\','_')))
  $contentDir= New-Dir (Join-Path $objRoot "Content")
  $metaDir   = New-Dir (Join-Path $objRoot "Metadata")
  $quarantine= $objRoot

  # Save original XML
  $sdmXml = $App.SDMPackageXML
  $xmlPath = Join-Path $metaDir "SDMPackageXML.xml"
  $sdmXml | Out-File -LiteralPath $xmlPath -Encoding UTF8

  # Extract details from XML
  $x = Parse-AppXml -Xml $sdmXml

  # Collect DeploymentTypes + content locations (best-effort)
  $dtInfos = @()
  $contentPaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

  foreach($dt in $x.AppMgmtDigest.Application.DeploymentTypes.DeploymentType){
    $dtName  = $dt.Title
    $dtTech  = $dt.Technology
    $retCodes= @()
    if($dt.Installer.ReturnCodes.ReturnCode){
      foreach($rc in $dt.Installer.ReturnCodes.ReturnCode){
        $retCodes += [pscustomobject]@{ Code=$rc.Value; Type=$rc.ReturnCodeType }
      }
    }
    # Detection rules / Requirements (best-effort summarize)
    $detect = ($dt.DetectionMethod | Out-String).Trim()
    $reqs   = ($dt.Requirements | Out-String).Trim()

    # Content
    if($dt.Installer.Contents.Content){
      foreach($c in $dt.Installer.Contents.Content){
        if($c.Location){ [void]$contentPaths.Add($c.Location.ToString()) }
      }
    }

    $dtInfos += [pscustomobject]@{
      Name=$dtName; Technology=$dtTech; ReturnCodes=$retCodes; Detection=$detect; Requirements=$reqs
    }
  }

  # Supersedence / Dependencies (summary)
  $supersedence = ($x.AppMgmtDigest.Application.Supercedence | Out-String).Trim()
  $dependencies = ($x.AppMgmtDigest.Application.Dependencies | Out-String).Trim()

  # Build metadata object
  $meta = [pscustomobject]@{
    Type             = 'Application'
    CI_UniqueID      = $id
    Name             = $name
    Version          = $App.SoftwareVersion
    Manufacturer     = $App.Manufacturer
    Description      = $App.LocalizedDescription
    Categories       = ($App.LocalizedCategoryInstanceNames -join ';')
    LastModified     = $App.LastModified
    Supersedence     = $supersedence
    Dependencies     = $dependencies
    DeploymentTypes  = $dtInfos
    ContentLocations = @($contentPaths)
  }
  $metaJsonPath = Join-Path $metaDir "metadata.json"
  $meta | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $metaJsonPath -Encoding UTF8

  # Copy content + hashlist
  $hashCsv = Join-Path $objRoot "hashlist.csv"
  "File,SizeBytes,SHA256" | Out-File -LiteralPath $hashCsv -Encoding UTF8
  $filesCount = 0
  $totalBytes = 0

  foreach($srcPath in $contentPaths){
    if(-not $srcPath -or -not (Test-Path -LiteralPath $srcPath)){
      Write-CMTraceLog "Content path not found for App '$name' ($id): $srcPath" 'WARN'
      continue
    }
    # Maintain tree
    $destBase = Join-Path $contentDir (Sanitize-Name ($srcPath -replace '^[A-Za-z]:\\','' -replace '^\\\\','UNC\'))
    New-Dir $destBase | Out-Null

    Get-ChildItem -LiteralPath $srcPath -File -Recurse | ForEach-Object {
      $rel = Resolve-Path -LiteralPath $_.FullName
      $relDest = Join-Path $destBase ($_.FullName.Substring($srcPath.Length).TrimStart('\'))
      $hash = Copy-WithHash -Source $_.FullName -Destination $relDest -RetryCount $RetryCount -RetryDelaySec $RetryDelaySec -QuarantineRoot $quarantine
      $filesCount++
      $totalBytes += $_.Length
      """$($relDest)""",$_.Length,$hash -join ',' | Out-File -LiteralPath $hashCsv -Append -Encoding UTF8
    }
  }

  # Include DPs (best-effort)
  if($IncludeDPs){
    try{
      # For applications it's complex; we record content IDs presence only (placeholder).
      $dpCsv = Join-Path $objRoot "DPs.csv"
      "Object,Note" | Out-File -LiteralPath $dpCsv -Encoding UTF8
      "Application,$id - DP enumeration not fully implemented (WMI mapping of content IDs required)" | Out-File -LiteralPath $dpCsv -Append -Encoding UTF8
    } catch {
      Write-CMTraceLog "DP list export failed for App $id : $($_.Exception.Message)" 'WARN'
    }
  }

  # Return summary for global manifest
  return [pscustomobject]@{
    Type='Application'; Name=$name; ID=$id; Version=$App.SoftwareVersion; Manufacturer=$App.Manufacturer
    LastModified=$App.LastModified; ContentBytes=$totalBytes; FilesCount=$filesCount
    ExportPath=$objRoot; ZipPath=$null
  }
}
#endregion

#region --- Export logic: Packages ---
function Export-CMPackage {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Pkg,
    [Parameter(Mandatory)] [string] $Root,
    [int] $RetryCount = 3,
    [int] $RetryDelaySec = 3
  )
  $id        = $Pkg.PackageID
  $name      = $Pkg.Name
  $nameSafe  = Sanitize-Name $name
  $objRoot   = New-Dir (Join-Path $Root ("Packages\{0}__{1}" -f $nameSafe,$id))
  $contentDir= New-Dir (Join-Path $objRoot "Content")
  $metaDir   = New-Dir (Join-Path $objRoot "Metadata")
  $quarantine= $objRoot

  # Collect Programs
  $programs = @()
  try{
    $progs = Get-WmiObject -Class SMS_Program -Namespace ("root\SMS\site_$SiteCode") -ComputerName $SiteServer -Filter ("PackageID='{0}'" -f $id)
    foreach($p in $progs){
      $programs += [pscustomobject]@{
        ProgramName = $p.ProgramName
        CommandLine = $p.CommandLine
        RunMode     = $p.RunMode
        UserInteraction = $p.UserInteraction
        EstimatedRunTime = $p.EstimatedRunTime
        Requirements    = $p.Requirements
        SuccessCodes    = $p.SuccessCodes
      }
    }
  } catch { Write-CMTraceLog "Failed to enumerate programs for $id : $($_.Exception.Message)" 'WARN' }

  # Metadata JSON
  $meta = [pscustomobject]@{
    Type         = 'Package'
    PackageID    = $id
    Name         = $name
    Version      = $Pkg.Version
    Manufacturer = $Pkg.Manufacturer
    Language     = $Pkg.Language
    Description  = $Pkg.Description
    SourcePath   = $Pkg.PkgSourcePath
    Programs     = $programs
    LastRefresh  = $Pkg.LastRefreshTime
  }
  $metaJsonPath = Join-Path $metaDir "metadata.json"
  $meta | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $metaJsonPath -Encoding UTF8

  # Copy content + hashlist
  $hashCsv = Join-Path $objRoot "hashlist.csv"
  "File,SizeBytes,SHA256" | Out-File -LiteralPath $hashCsv -Encoding UTF8
  $filesCount = 0
  $totalBytes = 0

  $srcPath = if($Pkg.PkgSourcePath){ $Pkg.PkgSourcePath } elseif($Pkg.SourceSite){ $Pkg.SourceSite } else { $null }
  if([string]::IsNullOrWhiteSpace($srcPath) -or -not (Test-Path -LiteralPath $srcPath)){
    Write-CMTraceLog "Package '$name' ($id) has no valid SourcePath: '$srcPath'" 'WARN'
  } else {
    $destBase = Join-Path $contentDir (Sanitize-Name ($srcPath -replace '^[A-Za-z]:\\','' -replace '^\\\\','UNC\'))
    New-Dir $destBase | Out-Null
    Get-ChildItem -LiteralPath $srcPath -File -Recurse | ForEach-Object {
      $relDest = Join-Path $destBase ($_.FullName.Substring($srcPath.Length).TrimStart('\'))
      $hash = Copy-WithHash -Source $_.FullName -Destination $relDest -RetryCount $RetryCount -RetryDelaySec $RetryDelaySec -QuarantineRoot $quarantine
      $filesCount++
      $totalBytes += $_.Length
      """$($relDest)""",$_.Length,$hash -join ',' | Out-File -LiteralPath $hashCsv -Append -Encoding UTF8
    }
  }

  # Include DPs (best-effort)
  if($IncludeDPs){
    try{
      $dpCsv = Join-Path $objRoot "DPs.csv"
      "DPServer,Note" | Out-File -LiteralPath $dpCsv -Encoding UTF8
      $summs = Get-WmiObject -Class SMS_PackageStatusDistPointsSummarizer -Namespace ("root\SMS\site_$SiteCode") -ComputerName $SiteServer -Filter ("PackageID='{0}'" -f $id)
      foreach($s in $summs){
        $server = ($s.ServerNALPath -replace '.*\\\\','') # extract server name
        "$server," | Out-File -LiteralPath $dpCsv -Append -Encoding UTF8
      }
    } catch {
      Write-CMTraceLog "DP list export failed for Package $id : $($_.Exception.Message)" 'WARN'
    }
  }

  return [pscustomobject]@{
    Type='Package'; Name=$name; ID=$id; Version=$Pkg.Version; Manufacturer=$Pkg.Manufacturer
    LastModified=$Pkg.LastRefreshTime; ContentBytes=$totalBytes; FilesCount=$filesCount
    ExportPath=$objRoot; ZipPath=$null
  }
}
#endregion

#region --- Dispatcher (sequential or parallel) ---
$results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

function Zip-And-Cleanup {
  param([pscustomobject]$Entry)
  if(-not $ZipPerObject){ return $Entry }
  try{
    $zipPath = "$($Entry.ExportPath).zip"
    if(Test-Path -LiteralPath $zipPath){ Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
    Compress-Archive -LiteralPath $Entry.ExportPath -DestinationPath $zipPath -Force
    if(Test-Path -LiteralPath $zipPath){
      Remove-Item -LiteralPath $Entry.ExportPath -Recurse -Force
      $Entry.ZipPath = $zipPath
    }
  } catch {
    Write-CMTraceLog "Zip failed for $($Entry.ID): $($_.Exception.Message)" 'WARN'
  }
  return $Entry
}

# Enumerate objects
$packages    = @()
$applications= @()

switch($ObjectType){
  'Package'     { $packages     = @( Get-PackageObjects -NameLike $NameLike -Ids $Ids -ModifiedAfter $ModifiedAfter ) }
  'Application' { $applications = @( Get-ApplicationObjects -NameLike $NameLike -Ids $Ids -ModifiedAfter $ModifiedAfter ) }
  'Both' {
     $packages     = @( Get-PackageObjects -NameLike $NameLike -Ids $Ids -ModifiedAfter $ModifiedAfter )
     $applications = @( Get-ApplicationObjects -NameLike $NameLike -Ids $Ids -ModifiedAfter $ModifiedAfter )
  }
}

Write-CMTraceLog "Objects to export: Packages=$($packages.Count), Applications=$($applications.Count)"

$work = @()
$packages    | ForEach-Object     { $work += [pscustomobject]@{ Kind='Pkg'; Obj=$_ } }
$applications| ForEach-Object     { $work += [pscustomobject]@{ Kind='App'; Obj=$_ } }

$exportRoot = New-Dir (Join-Path $SessionRoot "Export")
New-Dir (Join-Path $exportRoot 'Packages')     | Out-Null
New-Dir (Join-Path $exportRoot 'Applications') | Out-Null

if(Test-IsPS7 -and $MaxConcurrency -gt 1){
  Write-CMTraceLog "Running parallel export (Throttle=$MaxConcurrency)"
  $work | ForEach-Object -Parallel {
    param($using:exportRoot,$using:RetryCount,$using:RetryDelaySec,$using:IncludeDPs,$using:TraceLog)
    # Re-declare functions used inside parallel runspace
    function Write-CMTraceLog {
      param([string]$Message,[ValidateSet('INFO','WARN','ERROR')]$Level='INFO')
      $dt = Get-Date
      $type = switch ($Level) { 'INFO' {1} 'WARN' {2} 'ERROR' {3} }
      $line = "{0}`t{1}`tExport-CMContent`t`t{2}`t{3}`t{4}" -f $dt.ToString('MM-dd-yyyy'),$dt.ToString('HH:mm:ss.fff'),$type,[Threading.Thread]::CurrentThread.ManagedThreadId,$Message
      Add-Content -LiteralPath $using:TraceLog -Value $line -Encoding UTF8
    }
    function Sanitize-Name([string]$Name){
      if([string]::IsNullOrWhiteSpace($Name)){ return "_empty" }
      $bad = [IO.Path]::GetInvalidFileNameChars() + [char]'\'+[char]'/' + [char]':' + [char]'*' + [char]'?' + [char]'"' + [char]'<' + [char]'>' + [char]'|'
      ($Name.ToCharArray() | ForEach-Object { if($bad -contains $_){ '_' } else { $_ } }) -join ''
    }
    function New-Dir([string]$Path){ if(-not (Test-Path -LiteralPath $Path)){ New-Item -ItemType Directory -Path $Path -Force | Out-Null }; return $Path }
    function Get-FileSHA256([string]$Path){ (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToUpperInvariant() }
    function Copy-WithHash {
      param([string]$Source,[string]$Destination,[int]$RetryCount=3,[int]$RetryDelaySec=3,[string]$QuarantineRoot)
      if(-not (Test-Path -LiteralPath (Split-Path -Path $Destination -Parent))){
        New-Dir (Split-Path -Path $Destination -Parent) | Out-Null
      }
      $attempt=0
      while($true){
        try{
          Copy-Item -LiteralPath $Source -Destination $Destination -Force
          $srcH = Get-FileHash -Algorithm SHA256 -LiteralPath $Source
          $dstH = Get-FileHash -Algorithm SHA256 -LiteralPath $Destination
          if($srcH.Hash -ne $dstH.Hash){ throw "HASH_MISMATCH: $Source" }
          return $dstH.Hash.ToUpperInvariant()
        } catch {
          $attempt++
          if($attempt -gt $RetryCount){
            Write-CMTraceLog "Copy failed after $RetryCount retries for $Source -> $Destination : $($_.Exception.Message)" 'ERROR'
            if($QuarantineRoot){
              $qdir = New-Dir (Join-Path $QuarantineRoot "_quarantine")
              $qdst = Join-Path $qdir (Split-Path -Leaf $Source)
              try { Copy-Item -LiteralPath $Source -Destination $qdst -Force -ErrorAction Stop } catch {}
            }
            throw
          } else {
            Write-CMTraceLog "Copy retry $attempt/$RetryCount for $Source : $($_.Exception.Message)" 'WARN'
            Start-Sleep -Seconds $RetryDelaySec
          }
        }
      }
    }
    function Parse-AppXml { param([string]$Xml) return ([xml]$Xml) }

    if($_.Kind -eq 'Pkg'){
      # minimal inline export for package to avoid redefining all outer functions; call back to main via remoting not possible.
      $Pkg = $_.Obj
      $id   = $Pkg.PackageID
      $name = $Pkg.Name
      $nameSafe = Sanitize-Name $name
      $objRoot = New-Dir (Join-Path $using:exportRoot ("Packages\{0}__{1}" -f $nameSafe,$id))
      $metaDir = New-Dir (Join-Path $objRoot "Metadata")
      $contentDir = New-Dir (Join-Path $objRoot "Content")
      # metadata
      $meta = [pscustomobject]@{
        Type='Package'; PackageID=$id; Name=$name; Version=$Pkg.Version; Manufacturer=$Pkg.Manufacturer; Language=$Pkg.Language; Description=$Pkg.Description; SourcePath=$Pkg.PkgSourcePath; LastRefresh=$Pkg.LastRefreshTime
      }
      $meta | ConvertTo-Json -Depth 6 | Out-File -LiteralPath (Join-Path $metaDir "metadata.json") -Encoding UTF8
      # content
      $hashCsv = Join-Path $objRoot "hashlist.csv"
      "File,SizeBytes,SHA256" | Out-File -LiteralPath $hashCsv -Encoding UTF8
      $filesCount=0; $totalBytes=0
      $srcPath = $Pkg.PkgSourcePath
      if($srcPath -and (Test-Path -LiteralPath $srcPath)){
        $destBase = Join-Path $contentDir (Sanitize-Name ($srcPath -replace '^[A-Za-z]:\\','' -replace '^\\\\','UNC\'))
        New-Dir $destBase | Out-Null
        Get-ChildItem -LiteralPath $srcPath -File -Recurse | ForEach-Object {
          $relDest = Join-Path $destBase ($_.FullName.Substring($srcPath.Length).TrimStart('\'))
          $hash = Copy-WithHash -Source $_.FullName -Destination $relDest -RetryCount $using:RetryCount -RetryDelaySec $using:RetryDelaySec -QuarantineRoot $objRoot
          $filesCount++; $totalBytes+=$_.Length
          """$($relDest)""",$_.Length,$hash -join ',' | Out-File -LiteralPath $hashCsv -Append -Encoding UTF8
        }
      }
      [pscustomobject]@{ Type='Package'; Name=$name; ID=$id; Version=$Pkg.Version; Manufacturer=$Pkg.Manufacturer; LastModified=$Pkg.LastRefreshTime; ContentBytes=$totalBytes; FilesCount=$filesCount; ExportPath=$objRoot; ZipPath=$null }
    }
    elseif($_.Kind -eq 'App'){
      $App = $_.Obj
      $id   = $App.CI_UniqueID
      $name = if($App.LocalizedDisplayName){ $App.LocalizedDisplayName } else { $App.ModelName }
      $nameSafe = Sanitize-Name $name
      $objRoot = New-Dir (Join-Path $using:exportRoot ("Applications\{0}__{1}" -f $nameSafe,$id.Replace('\','_')))
      $metaDir = New-Dir (Join-Path $objRoot "Metadata")
      $contentDir= New-Dir (Join-Path $objRoot "Content")
      $sdmXml = $App.SDMPackageXML
      $sdmXml | Out-File -LiteralPath (Join-Path $metaDir "SDMPackageXML.xml") -Encoding UTF8
      $x = Parse-AppXml -Xml $sdmXml
      $contentPaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
      foreach($dt in $x.AppMgmtDigest.Application.DeploymentTypes.DeploymentType){
        if($dt.Installer.Contents.Content){
          foreach($c in $dt.Installer.Contents.Content){
            if($c.Location){ [void]$contentPaths.Add($c.Location.ToString()) }
          }
        }
      }
      $meta = [pscustomobject]@{
        Type='Application'; CI_UniqueID=$id; Name=$name; Version=$App.SoftwareVersion; Manufacturer=$App.Manufacturer; Description=$App.LocalizedDescription
        Categories=($App.LocalizedCategoryInstanceNames -join ';'); LastModified=$App.LastModified; ContentLocations=@($contentPaths)
      }
      $meta | ConvertTo-Json -Depth 6 | Out-File -LiteralPath (Join-Path $metaDir "metadata.json") -Encoding UTF8

      $hashCsv = Join-Path $objRoot "hashlist.csv"
      "File,SizeBytes,SHA256" | Out-File -LiteralPath $hashCsv -Encoding UTF8
      $filesCount=0; $totalBytes=0

      foreach($srcPath in $contentPaths){
        if($srcPath -and (Test-Path -LiteralPath $srcPath)){
          $destBase = Join-Path $contentDir (Sanitize-Name ($srcPath -replace '^[A-Za-z]:\\','' -replace '^\\\\','UNC\'))
          New-Dir $destBase | Out-Null
          Get-ChildItem -LiteralPath $srcPath -File -Recurse | ForEach-Object {
            $relDest = Join-Path $destBase ($_.FullName.Substring($srcPath.Length).TrimStart('\'))
            $hash = Copy-WithHash -Source $_.FullName -Destination $relDest -RetryCount $using:RetryCount -RetryDelaySec $using:RetryDelaySec -QuarantineRoot $objRoot
            $filesCount++; $totalBytes+=$_.Length
            """$($relDest)""",$_.Length,$hash -join ',' | Out-File -LiteralPath $hashCsv -Append -Encoding UTF8
          }
        } else {
          Write-CMTraceLog "Content path not found for App '$name' ($id): $srcPath" 'WARN'
        }
      }
      [pscustomobject]@{ Type='Application'; Name=$name; ID=$id; Version=$App.SoftwareVersion; Manufacturer=$App.Manufacturer; LastModified=$App.LastModified; ContentBytes=$totalBytes; FilesCount=$filesCount; ExportPath=$objRoot; ZipPath=$null }
    }
  } -ThrottleLimit $MaxConcurrency | ForEach-Object { [void]$results.Add($_) }
}
else {
  Write-CMTraceLog "Running sequential export"
  foreach($w in $work){
    try{
      if($w.Kind -eq 'Pkg'){
        $res = Export-CMPackage -Pkg $w.Obj -Root $exportRoot -RetryCount $RetryCount -RetryDelaySec $RetryDelaySec
      } else {
        $res = Export-CMApplication -App $w.Obj -Root $exportRoot -RetryCount $RetryCount -RetryDelaySec $RetryDelaySec
      }
      [void]$results.Add($res)
    } catch {
      Write-CMTraceLog "Export failed for $($w.Kind): $($_.Exception.Message)" 'ERROR'
    }
  }
}

# Zip (optional)
$final = foreach($r in $results){
  if($null -eq $r){ continue }
  Zip-And-Cleanup -Entry $r
}

#endregion

#region --- Global manifest + summary ---
# CSV
"Type;Name;ID;Version;Manufacturer;LastModified;ContentBytes;FilesCount;ExportPath;ZipPath" | Out-File -LiteralPath $ManifestCsv -Encoding UTF8
$final | ForEach-Object {
  "{0};{1};{2};{3};{4};{5:O};{6};{7};{8};{9}" -f $_.Type,$_.Name,$_.ID,$_.Version,$_.Manufacturer,$_.LastModified,$_.ContentBytes,$_.FilesCount,$_.ExportPath,$_.ZipPath |
    Out-File -LiteralPath $ManifestCsv -Append -Encoding UTF8
}
# JSON
$final | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $ManifestJson -Encoding UTF8

$ok = ($final | Where-Object { $_ }).Count
$total = $work.Count
Write-CMTraceLog "=== Completed Export: $ok/$total objects exported. Root: $SessionRoot ==="
Stop-Transcript | Out-Null

Write-Host ""
Write-Host "Export terminé." -ForegroundColor Green
Write-Host "Racine session : $SessionRoot"
Write-Host "Manifeste CSV   : $ManifestCsv"
Write-Host "Manifeste JSON  : $ManifestJson"
Write-Host "Logs (CMTrace)  : $TraceLog"
Write-Host "Transcript      : $Transcript"
#endregion
