  #requires -Version 5.1

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  Add-Type -AssemblyName System.Web

  $script:SysAdminClientVersion = '1.1.0'

  $script:PackageService = [PSCustomObject]@{
      Endpoint   = if ($env:SYSADMIN_GAS_URL) { $env:SYSADMIN_GAS_URL.Trim() } else { '' }
      ApiKey     = if ($env:SYSADMIN_GAS_KEY) { $env:SYSADMIN_GAS_KEY.Trim() } else { '' }
      TimeoutSec = 15
  }

  if ($env:SYSADMIN_GAS_TIMEOUT) {
      $timeoutValue = 0
      if ([int]::TryParse($env:SYSADMIN_GAS_TIMEOUT, [ref]$timeoutValue)) {
          $script:PackageService.TimeoutSec = $timeoutValue
      }
  }

  $script:InstallPackagesCache = $null
  $script:InstallPackagesSource = 'Default'
  $script:InstallPackagesLastSync = $null
  $script:InstallPackagesLastError = $null
  $script:InstallPackagesHasAnnounced = $false

  [System.Windows.Forms.Application]::EnableVisualStyles()

  function Set-GridData {
      param(
          [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
          [Parameter()][System.Collections.IEnumerable]$Data
      )
      $Grid.DataSource = $null
      $bindingList = New-Object System.Collections.ArrayList
      if ($Data) {
          foreach ($row in $Data) {
              [void]$bindingList.Add($row)
          }
      }
      $Grid.DataSource = $bindingList
  }

  function Convert-NetworkStatus {
      param([int]$Status)
      switch ($Status) {
          0 { "Disconnected" }
          1 { "Connecting" }
          2 { "Connected" }
          3 { "Disconnecting" }
          4 { "Hardware Not Present" }
          5 { "Hardware Disabled" }
          6 { "Hardware Malfunction" }
          7 { "Media Disconnected" }
          8 { "Authenticating" }
          9 { "Auth Succeeded" }
          10 { "Auth Failed" }
          11 { "Invalid Address" }
          12 { "Credentials Required" }
          default { "Unknown" }
      }
  }

  function Get-HardwareData {
      [CmdletBinding()]
      param([Parameter(Mandatory)][ValidateSet("CPU","Board","GPU","Memory","Disk","Network")]$Category)

      switch ($Category) {
          "CPU" {
              try {
                  Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object `
                      @{Name='Name';Expression={$_.Name.Trim()}},
                      Manufacturer,
                      SocketDesignation,
                      NumberOfCores,
                      NumberOfLogicalProcessors,
                      @{Name='MaxClockGHz';Expression={ "{0:N2}" -f ($_.MaxClockSpeed / 1000) }},
                      @{Name='L3CacheMB';Expression={ if ($_.L3CacheSize) { "{0:N0}" -f ($_.L3CacheSize / 1024) } else
  { "N/A" } }}
              } catch {
                  ,([PSCustomObject]@{Message="Unable to read CPU information."; Detail=$_.Exception.Message})
              }
          }
          "Board" {
              try {
                  Get-CimInstance Win32_BaseBoard -ErrorAction Stop | Select-Object `
                      Manufacturer,
                      Product,
                      SerialNumber,
                      Version,
                      @{Name='PoweredOn';Expression={$_.PoweredOn}}
              } catch {
                  ,([PSCustomObject]@{Message="Unable to read baseboard information."; Detail=$_.Exception.Message})
              }
          }
          "GPU" {
              try {
                  Get-CimInstance Win32_VideoController -ErrorAction Stop | Select-Object `
                      Name,
                      DriverVersion,
                      VideoProcessor,
                      CurrentRefreshRate,
                      @{Name='AdapterRAM_MB';Expression={
                          if ($_.AdapterRAM -and $_.AdapterRAM -gt 0) { "{0:N0}" -f ($_.AdapterRAM / 1MB) } else { "N/
  A" }
                      }},
                      @{Name='Status';Expression={$_.Status}}
              } catch {
                  ,([PSCustomObject]@{Message="Unable to read GPU information."; Detail=$_.Exception.Message})
              }
          }
          "Memory" {
              try {
                  Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop | Select-Object `
                      BankLabel,
                      DeviceLocator,
                      @{Name='CapacityGB';Expression={ "{0:N2}" -f (($_.Capacity)/1GB) }},
                      Speed,
                      Manufacturer,
                      SerialNumber
              } catch {
                  ,([PSCustomObject]@{Message="Unable to read memory information."; Detail=$_.Exception.Message})
              }
          }
          "Disk" {
              try {
                  Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop | Select-Object `
                      DeviceID,
                      VolumeName,
                      FileSystem,
                      @{Name='SizeGB';Expression={ "{0:N1}" -f (($_.Size)/1GB) }},
                      @{Name='FreeGB';Expression={ "{0:N1}" -f (($_.FreeSpace)/1GB) }},
                      @{Name='FreePercent';Expression={
                          if ($_.Size) { "{0:N1}%" -f (($_.FreeSpace/$_.Size)*100) } else { "N/A" }
                      }}
              } catch {
                  ,([PSCustomObject]@{Message="Unable to read disk information."; Detail=$_.Exception.Message})
              }
          }
          "Network" {
              try {
                  $adapters = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -ne 'Disabled' }
                  if ($adapters) {
                      return $adapters | Select-Object `
                          Name,
                          InterfaceDescription,
                          Status,
                          @{Name='LinkSpeed_Mbps';Expression={ if ($_.LinkSpeed) { "{0:N1}" -f ($_.LinkSpeed / 1e6) }
  else { "N/A" } }},
                          MacAddress
                  }
              } catch { }
              try {
                  Get-CimInstance Win32_NetworkAdapter -Filter "PhysicalAdapter=True" -ErrorAction Stop |
                      Where-Object { $_.NetEnabled -eq $true } |
                      Select-Object `
                          Name,
                          @{Name='InterfaceDescription';Expression={$_.Description}},
                          @{Name='Status';Expression={ Convert-NetworkStatus $_.NetConnectionStatus }},
                          @{Name='LinkSpeed_Mbps';Expression={ if ($_.Speed) { "{0:N1}" -f ($_.Speed / 1e6) } else { "N/
  A" } }},
                          MACAddress
              } catch {
                  ,([PSCustomObject]@{Message="Unable to read network information."; Detail=$_.Exception.Message})
              }
          }
      }
  }

  function Get-ZipEntryText {
      param(
          [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Archive,
          [Parameter(Mandatory)][string]$EntryName
      )
      $entry = $Archive.GetEntry($EntryName)
      if (-not $entry) { return $null }
      $stream = $entry.Open()
      $reader = New-Object System.IO.StreamReader($stream)
      try {
          $reader.ReadToEnd()
      } finally {
          $reader.Dispose()
          $stream.Dispose()
      }
  }

  function Get-DocxText {
      param([Parameter(Mandatory)][string]$Path)
      $fs = $null
      $archive = $null
      try {
          $fs = [System.IO.File]::Open($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
          $archive = New-Object System.IO.Compression.ZipArchive($fs,[System.IO.Compression.ZipArchiveMode]::Read)
          $builder = New-Object System.Text.StringBuilder
          foreach ($entry in $archive.Entries) {
              if ($entry.FullName -like "word/*.xml" -and $entry.FullName -notlike "word/_rels/*") {
                  $xml = Get-ZipEntryText -Archive $archive -EntryName $entry.FullName
                  if ($xml) {
                      try {
                          $doc = New-Object System.Xml.XmlDocument
                          $doc.PreserveWhitespace = $false
                          $doc.LoadXml($xml)
                          $textNodes = $doc.SelectNodes("//*[local-name()='t']")
                          if ($textNodes -and $textNodes.Count -gt 0) {
                              foreach ($node in $textNodes) {
                                  $value = $node.InnerText
                                  if (-not [string]::IsNullOrWhiteSpace($value)) {
                                      [void]$builder.AppendLine($value)
                                  }
                              }
                          } else {
                              $text = $doc.InnerText
                              if (-not [string]::IsNullOrWhiteSpace($text)) {
                                  [void]$builder.AppendLine($text)
                              }
                          }
                      } catch {
                          $fallback = ($xml -replace '<[^>]+>',' ')
                          if (-not [string]::IsNullOrWhiteSpace($fallback)) {
                              [void]$builder.AppendLine($fallback)
                          }
                      }
                  }
              }
          }
          $result = $builder.ToString()
          if ([string]::IsNullOrWhiteSpace($result)) { return $null }
          $result
      } catch {
          $null
      } finally {
          if ($archive) { $archive.Dispose() }
          if ($fs) { $fs.Dispose() }
      }
  }


  function Get-XlsxText {
      param([Parameter(Mandatory)][string]$Path)
      $fs = $null
      $archive = $null
      try {
          $fs = [System.IO.File]::Open($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
          $archive = New-Object System.IO.Compression.ZipArchive($fs,[System.IO.Compression.ZipArchiveMode]::Read)
          $builder = New-Object System.Text.StringBuilder
          $sharedStrings = New-Object System.Collections.Generic.List[string]
          $shared = Get-ZipEntryText -Archive $archive -EntryName "xl/sharedStrings.xml"
          if ($shared) {
              try {
                  $sharedDoc = New-Object System.Xml.XmlDocument
                  $sharedDoc.LoadXml($shared)
                  foreach ($si in $sharedDoc.GetElementsByTagName("si")) {
                      $value = $si.InnerText
                      if ($null -ne $value) {
                          [void]$sharedStrings.Add($value)
                      } else {
                          [void]$sharedStrings.Add("")
                      }
                  }
              } catch {
                  $fallbackShared = ($shared -replace '<[^>]+>',' ')
                  if (-not [string]::IsNullOrWhiteSpace($fallbackShared)) {
                      [void]$builder.AppendLine($fallbackShared)
                  }
              }
          }
          foreach ($entry in $archive.Entries) {
              if ($entry.FullName -like "xl/worksheets/*.xml") {
                  $xml = Get-ZipEntryText -Archive $archive -EntryName $entry.FullName
                  if ($xml) {
                      try {
                          $doc = New-Object System.Xml.XmlDocument
                          $doc.LoadXml($xml)
                          foreach ($cell in $doc.SelectNodes("//*[local-name()='c']")) {
                              $typeAttr = $cell.Attributes["t"]
                              $cellText = $null
                              if ($typeAttr -and $typeAttr.Value -eq "s") {
                                  $valueNode = $cell.SelectSingleNode("*[local-name()='v']")
                                  if ($valueNode) {
                                      try {
                                          $index = [int]$valueNode.InnerText
                                          if ($index -ge 0 -and $index -lt $sharedStrings.Count) {
                                              $cellText = $sharedStrings[$index]
                                          }
                                      } catch { }
                                  }
                              } elseif ($typeAttr -and $typeAttr.Value -eq "inlineStr") {
                                  $inline = $cell.SelectSingleNode("*[local-name()='is']")
                                  if ($inline) {
                                      $cellText = $inline.InnerText
                                  }
                              } else {
                                  $valueNode = $cell.SelectSingleNode("*[local-name()='v']")
                                  if ($valueNode) {
                                      $cellText = $valueNode.InnerText
                                  }
                              }
                              if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                                  [void]$builder.AppendLine($cellText)
                              }
                          }
                      } catch {
                          $fallbackSheet = ($xml -replace '<[^>]+>',' ')
                          if (-not [string]::IsNullOrWhiteSpace($fallbackSheet)) {
                              [void]$builder.AppendLine($fallbackSheet)
                          }
                      }
                  }
              }
          }
          $result = $builder.ToString()
          if ([string]::IsNullOrWhiteSpace($result)) { return $null }
          $result
      } catch {
          $null
      } finally {
          if ($archive) { $archive.Dispose() }
          if ($fs) { $fs.Dispose() }
      }
  }

  function Convert-PdfEscapedString {
      param([string]$Value)
      if ($null -eq $Value) { return $null }
      $builder = New-Object System.Text.StringBuilder
      $length = $Value.Length
      for ($i = 0; $i -lt $length; $i++) {
          $char = $Value[$i]
          if ($char -eq '\') {
              $i++
              if ($i -ge $length) { break }
              $escape = $Value[$i]
              if ($escape -eq 'n') { [void]$builder.Append([char]10); continue }
              if ($escape -eq 'r') { [void]$builder.Append([char]13); continue }
              if ($escape -eq 't') { [void]$builder.Append([char]9); continue }
              if ($escape -eq 'b') { [void]$builder.Append([char]8); continue }
              if ($escape -eq 'f') { [void]$builder.Append([char]12); continue }
              if ($escape -eq '(') { [void]$builder.Append('('); continue }
              if ($escape -eq ')') { [void]$builder.Append(')'); continue }
              if ($escape -eq '\') { [void]$builder.Append('\'); continue }
              if ($escape -eq [char]13 -or $escape -eq [char]10) {
                  if ($escape -eq [char]13 -and $i + 1 -lt $length -and $Value[$i + 1] -eq [char]10) { $i++ }
                  continue
              }
              if ($escape -ge '0' -and $escape -le '7') {
                  $octal = "" + $escape
                  $max = [System.Math]::Min(2, $length - $i - 1)
                  for ($j = 0; $j -lt $max; $j++) {
                      $nextChar = $Value[$i + 1]
                      if ($nextChar -ge '0' -and $nextChar -le '7') {
                          $i++
                          $octal += $nextChar
                      } else {
                          break
                      }
                  }
                  try {
                      $code = [Convert]::ToInt32($octal,8)
                      [void]$builder.Append([char]$code)
                  } catch {
                      [void]$builder.Append($octal)
                  }
                  continue
              }
              [void]$builder.Append($escape)
              continue
          }
          [void]$builder.Append($char)
      }
      $builder.ToString()
  }

  function Extract-PdfStrings {
      param([string]$Source)
      $fragments = New-Object System.Collections.Generic.List[string]
      if ([string]::IsNullOrWhiteSpace($Source)) { return $fragments }
      $length = $Source.Length
      $builder = $null
      $depth = 0
      for ($i = 0; $i -lt $length; $i++) {
          $char = $Source[$i]
          if ($null -eq $builder) {
              if ($char -eq '(') {
                  $builder = New-Object System.Text.StringBuilder
                  $depth = 1
              }
              continue
          }
          if ($char -eq '\') {
              if ($i + 1 -lt $length) {
                  [void]$builder.Append($char)
                  $i++
                  [void]$builder.Append($Source[$i])
              }
              continue
          }
          if ($char -eq '(') {
              $depth++
              [void]$builder.Append($char)
              continue
          }
          if ($char -eq ')') {
              $depth--
              if ($depth -le 0) {
                  $raw = $builder.ToString()
                  $decoded = Convert-PdfEscapedString -Value $raw
                  if (-not [string]::IsNullOrWhiteSpace($decoded)) {
                      [void]$fragments.Add($decoded)
                  }
                  $builder = $null
                  $depth = 0
              } else {
                  [void]$builder.Append($char)
              }
              continue
          }
          [void]$builder.Append($char)
      }
      $fragments
  }

  function Invoke-PdfFlateDecode {
      param([byte[]]$Buffer)
      if (-not $Buffer -or $Buffer.Length -eq 0) { return $null }
      $memory = New-Object System.IO.MemoryStream
      $deflate = $null
      $reader = $null
      try {
          $memory.Write($Buffer,0,$Buffer.Length)
          $memory.Position = 0
          $deflate = New-Object System.IO.Compression.DeflateStream($memory,[System.IO.Compression.CompressionMode]::Decompress,$true)
          $reader = New-Object System.IO.StreamReader($deflate,[System.Text.Encoding]::UTF8,$true,1024,$true)
          $reader.ReadToEnd()
      } catch {
          $null
      } finally {
          if ($reader) { $reader.Dispose() }
          if ($deflate) { $deflate.Dispose() }
          $memory.Dispose()
      }
  }

  function Get-PdfText {
      param([Parameter(Mandatory)][string]$Path)
      try {
          $bytes = [System.IO.File]::ReadAllBytes($Path)
          $latin = [System.Text.Encoding]::GetEncoding("ISO-8859-1")
          $builder = New-Object System.Text.StringBuilder
          $rawText = $latin.GetString($bytes)
          foreach ($fragment in (Extract-PdfStrings -Source $rawText)) {
              if (-not [string]::IsNullOrWhiteSpace($fragment)) {
                  [void]$builder.AppendLine($fragment)
              }
          }
          $streamRegex = New-Object System.Text.RegularExpressions.Regex("stream[
]+(?<data>.*?)[
]+endstream",[System.Text.RegularExpressions.RegexOptions]::Singleline)
          foreach ($match in $streamRegex.Matches($rawText)) {
              $dataGroup = $match.Groups['data']
              if ($dataGroup.Length -le 0) { continue }
              $shouldInflate = $false
              $prefixStart = [System.Math]::Max(0, $match.Index - 512)
              $prefixLength = $match.Index - $prefixStart
              if ($prefixLength -gt 0) {
                  $prefix = $rawText.Substring($prefixStart, $prefixLength)
                  if ($prefix -match "/FlateDecode") { $shouldInflate = $true }
              }
              $buffer = New-Object byte[]($dataGroup.Length)
              [System.Array]::Copy($bytes, $dataGroup.Index, $buffer, 0, $dataGroup.Length)
              $decoded = $null
              if ($shouldInflate -and $buffer.Length -gt 0) {
                  $decoded = Invoke-PdfFlateDecode -Buffer $buffer
                  if (-not $decoded -and $buffer.Length -gt 6) {
                      $trimmed = New-Object byte[]($buffer.Length - 6)
                      [System.Array]::Copy($buffer,2,$trimmed,0,$trimmed.Length)
                      $decoded = Invoke-PdfFlateDecode -Buffer $trimmed
                  }
              }
              if (-not $decoded) {
                  $decoded = $latin.GetString($buffer)
              }
              foreach ($fragment in (Extract-PdfStrings -Source $decoded)) {
                  if (-not [string]::IsNullOrWhiteSpace($fragment)) {
                      [void]$builder.AppendLine($fragment)
                  }
              }
          }
          $result = $builder.ToString()
          if ([string]::IsNullOrWhiteSpace($result)) { return $null }
          $result
      } catch {
          $null
      }
  }

  function Get-MatchSnippet {
      param(
          [Parameter(Mandatory)][string]$Source,
          [Parameter(Mandatory)][string]$Keyword,
          [int]$Context = 45
      )
      if ([string]::IsNullOrWhiteSpace($Source)) { return $null }
      $index = $Source.IndexOf($Keyword, [System.StringComparison]::OrdinalIgnoreCase)
      if ($index -lt 0) { return $null }
      $start = [System.Math]::Max(0, $index - $Context)
      $length = [System.Math]::Min($Context * 2 + $Keyword.Length, $Source.Length - $start)
      ($Source.Substring($start, $length) -replace '\s+',' ').Trim()
  }

  function Search-PlainText {
      param(
          [Parameter(Mandatory)][string]$Path,
          [Parameter(Mandatory)][string]$Keyword
      )
      try {
          $match = Select-String -Path $Path -Pattern $Keyword -SimpleMatch -List -ErrorAction Stop | Select-Object
  -First 1
          if ($match) { return $match.Line.Trim() }
      } catch {
          try {
              $content = Get-Content -Path $Path -Raw -ErrorAction Stop
              if ($content.IndexOf($Keyword,[System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                  return Get-MatchSnippet -Source $content -Keyword $Keyword
              }
          } catch { }
      }
      $null
  }

  function Get-FileContentMatch {
      param(
          [Parameter(Mandatory)][string]$Path,
          [Parameter(Mandatory)][string]$Keyword
      )
      $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
      switch ($ext) {
          ".txt" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".md" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".log" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".csv" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".json" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".xml" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".config" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".ini" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".ps1" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".psm1" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".cs" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".sql" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".bat" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".cmd" { Search-PlainText -Path $Path -Keyword $Keyword }
          ".docx" {
              $text = Get-DocxText -Path $Path
              if ($text) { Get-MatchSnippet -Source $text -Keyword $Keyword }
          }
          ".xlsx" {
              $text = Get-XlsxText -Path $Path
              if ($text) { Get-MatchSnippet -Source $text -Keyword $Keyword }
          }
          ".pdf" {
              $text = Get-PdfText -Path $Path
              if ($text) { Get-MatchSnippet -Source $text -Keyword $Keyword }
          }
          default { $null }
      }
  }

  function Invoke-FileSearch {
      param(
          [Parameter(Mandatory)][string]$Root,
          [Parameter(Mandatory)][string]$Keyword,
          [Parameter(Mandatory)][ValidateSet("Name","Content")]$Mode
      )

      $results = New-Object System.Collections.ArrayList
      if ($Mode -eq "Name") {
          $pattern = New-Object System.Management.Automation.WildcardPattern("*$Keyword*",
  [System.Management.Automation.WildcardOptions]::IgnoreCase)
          $files = Get-ChildItem -Path $Root -Recurse -File -ErrorAction SilentlyContinue
          foreach ($file in $files) {
              if ($script:SearchCancelled) { break }
              [System.Windows.Forms.Application]::DoEvents()
              if ($pattern.IsMatch($file.Name)) {
                  $item = [PSCustomObject]@{
                      Type  = "Name"
                      Path  = $file.FullName
                      Match = $file.Name
                  }
                  [void]$results.Add($item)
              }
          }
      } else {
          $allowed =
  @('.txt','.md','.log','.csv','.json','.xml','.config','.ini','.ps1','.psm1','.cs','.sql','.bat','.cmd','.docx','.xlsx'
  ,'.pdf')
          $files = Get-ChildItem -Path $Root -Recurse -File -ErrorAction SilentlyContinue
          foreach ($file in $files) {
              if ($script:SearchCancelled) { break }
              $extension = [System.IO.Path]::GetExtension($file.FullName).ToLowerInvariant()
              if ($allowed -contains $extension) {
                  [System.Windows.Forms.Application]::DoEvents()
                  $snippet = Get-FileContentMatch -Path $file.FullName -Keyword $Keyword
                  if ($snippet) {
                      $item = [PSCustomObject]@{
                          Type  = "Content"
                          Path  = $file.FullName
                          Match = $snippet
                      }
                      [void]$results.Add($item)
                  }
              }
          }
      }
      $results
  }

  function Get-StartupEntries {
      $targets = @(
          @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"; Scope="CurrentUser"},
          @{Path="HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"; Scope="LocalMachine"}
      )
      $list = New-Object System.Collections.ArrayList
      foreach ($target in $targets) {
          try {
              $key = Get-Item -Path $target.Path -ErrorAction Stop
              foreach ($name in $key.GetValueNames()) {
                  $value = $key.GetValue($name)
                  $item = [PSCustomObject]@{
                      Name         = $name
                      Command      = $value
                      Scope        = $target.Scope
                      RegistryPath = $target.Path
                  }
                  [void]$list.Add($item)
              }
          } catch { }
      }
      $list
  }

  function Add-StartupEntry {
      param(
          [Parameter(Mandatory)][string]$Name,
          [Parameter(Mandatory)][string]$Command,
          [Parameter(Mandatory)][ValidateSet("CurrentUser","LocalMachine")]$Scope
      )
      $path = if ($Scope -eq "CurrentUser") {
          "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
      } else {
          "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
      }
      Set-ItemProperty -Path $path -Name $Name -Value $Command -Force -ErrorAction Stop
  }

  function Remove-StartupEntry {
      param(
          [Parameter(Mandatory)][string]$Name,
          [Parameter(Mandatory)][string]$RegistryPath
      )
      Remove-ItemProperty -Path $RegistryPath -Name $Name -ErrorAction Stop
  }

  function Download-File {
      param(
          [Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][string]$Destination
      )
      if (Get-Command -Name Start-BitsTransfer -ErrorAction SilentlyContinue) {
          Start-BitsTransfer -Source $Url -Destination $Destination -ErrorAction Stop
      } else {
          Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -ErrorAction Stop
      }
  }

  function Get-DefaultInstallPackages {
      @(
          [PSCustomObject]@{
              Name        = "7-Zip 64-bit"
              Url         = "https://www.7-zip.org/a/7z2301-x64.exe"
              Description = "File archiver utility."
              Arguments   = "/S"
              Version     = ""
              Category    = "Utilities"
              Source      = "Default"
          },
          [PSCustomObject]@{
              Name        = "Visual Studio Code"
              Url         = "https://update.code.visualstudio.com/latest/win32-x64-user/stable"
              Description = "Lightweight source code editor."
              Arguments   = "/verysilent"
              Version     = ""
              Category    = "Development"
              Source      = "Default"
          },
          [PSCustomObject]@{
              Name        = "Git for Windows"
              Url         = "https://github.com/git-for-windows/git/releases/download/v2.45.1.windows.1/Git-2.45.1-64-bit.exe"
              Description = "Distributed version control."
              Arguments   = "/VERYSILENT"
              Version     = ""
              Category    = "Development"
              Source      = "Default"
          }
      )
  }

  function Invoke-PackageServiceRequest {
      param(
          [Parameter(Mandatory)][ValidateSet('GET','POST')][string]$Method,
          [Parameter(Mandatory)][string]$Action,
          [Parameter()][object]$Body
      )

      if (-not $script:PackageService.Endpoint) {
          throw [System.InvalidOperationException]::new("Package service endpoint is not configured.")
      }

      $uriBuilder = New-Object System.UriBuilder($script:PackageService.Endpoint)
      $query = [System.Web.HttpUtility]::ParseQueryString($uriBuilder.Query)
      $query["action"] = $Action
      if ($script:PackageService.ApiKey) {
          $query["key"] = $script:PackageService.ApiKey
      }
      $uriBuilder.Query = $query.ToString()
      $requestUri = $uriBuilder.Uri.AbsoluteUri

      if (-not ([Net.ServicePointManager]::SecurityProtocol.HasFlag([Net.SecurityProtocolType]::Tls12))) {
          [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
      }

      $arguments = @{
          Method      = $Method
          Uri         = $requestUri
          TimeoutSec  = $script:PackageService.TimeoutSec
          ErrorAction = 'Stop'
          Headers     = @{ Accept = 'application/json' }
      }

      if ($Method -eq 'POST' -and $null -ne $Body) {
          $arguments.Body = ($Body | ConvertTo-Json -Depth 6 -Compress)
          $arguments.ContentType = 'application/json'
      }

      return Invoke-RestMethod @arguments
  }

  function Get-RemoteInstallPackages {
      $response = Invoke-PackageServiceRequest -Method GET -Action 'packages'
      if (-not $response) {
          throw [System.InvalidOperationException]::new("Remote service returned no data.")
      }

      $packageData = $null
      if ($response.PSObject.Properties.Name -contains 'packages') {
          $packageData = $response.packages
      } elseif ($response -is [System.Collections.IEnumerable]) {
          $packageData = $response
      }

      $packages = @($packageData) | Where-Object { $_ }
      if (-not $packages -or $packages.Count -eq 0) {
          throw [System.InvalidOperationException]::new("Remote service returned an empty package list.")
      }

      $result = @()
      foreach ($pkg in $packages) {
          if (-not $pkg) { continue }
          $result += [PSCustomObject]@{
              Name        = [string]$pkg.name
              Url         = [string]$pkg.url
              Description = [string]$pkg.description
              Arguments   = if ($pkg.arguments) { [string]$pkg.arguments } else { '' }
              Version     = if ($pkg.version) { [string]$pkg.version } else { '' }
              Category    = if ($pkg.category) { [string]$pkg.category } else { '' }
              Source      = "Remote"
          }
      }

      return $result
  }

  function Get-InstallPackages {
      param([switch]$ForceRefresh)

      if (-not $ForceRefresh -and $script:InstallPackagesCache) {
          return $script:InstallPackagesCache
      }

      $packages = $null
      $source = 'Default'
      $errorMessage = $null

      if ($script:PackageService.Endpoint) {
          try {
              $remotePackages = Get-RemoteInstallPackages
              if ($remotePackages -and $remotePackages.Count -gt 0) {
                  $packages = $remotePackages
                  $source = 'Remote'
              } else {
                  $errorMessage = "Remote service returned no packages."
              }
          } catch {
              $errorMessage = $_.Exception.Message
          }
      }

      if (-not $packages) {
          $packages = Get-DefaultInstallPackages
      }

      $script:InstallPackagesCache = $packages
      $script:InstallPackagesSource = $source
      $script:InstallPackagesLastSync = Get-Date
      $script:InstallPackagesLastError = $errorMessage

      return $packages
  }

  function Send-InstallReport {
      param(
          [Parameter(Mandatory)][psobject]$Package,
          [Parameter(Mandatory)][string]$Status,
          [string]$Message,
          [double]$DurationSeconds,
          [string]$ErrorDetail
      )

      if (-not $script:PackageService.Endpoint) {
          return
      }

      $payload = @{
          packageName     = $Package.Name
          status          = $Status
          message         = $Message
          url             = $Package.Url
          arguments       = $Package.Arguments
          version         = $Package.Version
          category        = $Package.Category
          durationSeconds = if ($DurationSeconds) { [math]::Round($DurationSeconds, 2) } else { $null }
          user            = $env:USERNAME
          machine         = $env:COMPUTERNAME
          clientVersion   = $script:SysAdminClientVersion
          sentAt          = (Get-Date).ToUniversalTime().ToString("o")
      }

      if ($ErrorDetail) {
          $payload.error = $ErrorDetail
      }

      try {
          Invoke-PackageServiceRequest -Method POST -Action 'log' -Body $payload | Out-Null
      } catch {
          $statusControl = Get-Variable -Scope Script -Name installStatus -ErrorAction SilentlyContinue
          if ($statusControl) {
              Update-InstallStatus ("Failed to report install telemetry: {0}" -f $_.Exception.Message)
          }
      }
  }

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "System Management Console"
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.ClientSize = New-Object System.Drawing.Size(1100,650)
  $form.MinimumSize = New-Object System.Drawing.Size(1100,650)
  $form.Font = New-Object System.Drawing.Font("Segoe UI",9)

  $navPanel = New-Object System.Windows.Forms.Panel
  $navPanel.Dock = [System.Windows.Forms.DockStyle]::Left
  $navPanel.Width = 120
  $navPanel.Padding = New-Object System.Windows.Forms.Padding(10)
  $navPanel.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

  $contentPanel = New-Object System.Windows.Forms.Panel
  $contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $contentPanel.Padding = New-Object System.Windows.Forms.Padding(10)

  $form.Controls.Add($contentPanel)
  $form.Controls.Add($navPanel)

  $buttonSpecs = @(
      @{Text="1. System"; Key="System"},
      @{Text="2. Install"; Key="Install"},
      @{Text="3. Search"; Key="Search"},
      @{Text="4. Startup"; Key="Startup"}
  )
  $buttons = @()
  $buttonTop = 15
  foreach ($spec in $buttonSpecs) {
      $btn = New-Object System.Windows.Forms.Button
      $btn.Text = $spec.Text
      $btn.Width = 90
      $btn.Height = 40
      $btn.Left = 10
      $btn.Top = $buttonTop
      $btn.Tag = $spec.Key
      $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
      $btn.BackColor = [System.Drawing.Color]::White
      $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
      $navPanel.Controls.Add($btn)
      $buttons += $btn
      $buttonTop += 55
  }

  $systemPanel = New-Object System.Windows.Forms.Panel
  $systemPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $systemPanel.Visible = $false

  $systemLayout = New-Object System.Windows.Forms.TableLayoutPanel
  $systemLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $systemLayout.RowCount = 2
  $systemLayout.ColumnCount = 1
  $systemLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,50)))
  $systemLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
  $systemLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))

  $systemToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
  $systemToolbar.Dock = [System.Windows.Forms.DockStyle]::Fill
  $systemToolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $systemToolbar.WrapContents = $false
  $systemToolbar.Padding = New-Object System.Windows.Forms.Padding(0,10,0,10)

  $refreshHardwareButton = New-Object System.Windows.Forms.Button
  $refreshHardwareButton.Text = "Refresh Hardware"
  $refreshHardwareButton.AutoSize = $true
  $systemToolbar.Controls.Add($refreshHardwareButton)

  $hardwareTabs = New-Object System.Windows.Forms.TabControl
  $hardwareTabs.Dock = [System.Windows.Forms.DockStyle]::Fill

  $systemLayout.Controls.Add($systemToolbar,0,0)
  $systemLayout.Controls.Add($hardwareTabs,0,1)
  $systemPanel.Controls.Add($systemLayout)

  $script:hardwareGrids = @{}
  $hardwareCategories = @("CPU","Board","GPU","Memory","Disk","Network")
  foreach ($category in $hardwareCategories) {
      $tabPage = New-Object System.Windows.Forms.TabPage
      $tabPage.Text = $category
      $grid = New-Object System.Windows.Forms.DataGridView
      $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
      $grid.ReadOnly = $true
      $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
      $grid.RowHeadersVisible = $false
      $grid.AllowUserToAddRows = $false
      $grid.AllowUserToDeleteRows = $false
      $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
      $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
      $tabPage.Controls.Add($grid)
      [void]$hardwareTabs.TabPages.Add($tabPage)
      $script:hardwareGrids[$category] = $grid
  }

  function Update-HardwareView {
      foreach ($category in $hardwareCategories) {
          $grid = $script:hardwareGrids[$category]
          $data = @(Get-HardwareData -Category $category)
          Set-GridData -Grid $grid -Data $data
      }
  }
  $refreshHardwareButton.Add_Click({ Update-HardwareView })

  $installPanel = New-Object System.Windows.Forms.Panel
  $installPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $installPanel.Visible = $false

  $installLayout = New-Object System.Windows.Forms.TableLayoutPanel
  $installLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $installLayout.RowCount = 2
  $installLayout.ColumnCount = 1
  $installLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,60)))
  $installLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,40)))
  $installLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))

  $packageList = New-Object System.Windows.Forms.ListView
  $packageList.View = [System.Windows.Forms.View]::Details
  $packageList.FullRowSelect = $true
  $packageList.HideSelection = $false
  $packageList.MultiSelect = $false
  $packageList.Dock = [System.Windows.Forms.DockStyle]::Fill
  [void]$packageList.Columns.Add("Name",150)
  [void]$packageList.Columns.Add("Description",260)
  [void]$packageList.Columns.Add("URL",320)
  [void]$packageList.Columns.Add("Arguments",120)

  $installBottom = New-Object System.Windows.Forms.TableLayoutPanel
  $installBottom.Dock = [System.Windows.Forms.DockStyle]::Fill
  $installBottom.RowCount = 2
  $installBottom.ColumnCount = 1
  $installBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,50)))
  $installBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
  $installBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))

  $installButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $installButtonsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $installButtonsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $installButtonsPanel.WrapContents = $false
  $installButtonsPanel.Padding = New-Object System.Windows.Forms.Padding(5,10,5,5)

  $installButton = New-Object System.Windows.Forms.Button
  $installButton.Text = "Download && Install"
  $installButton.AutoSize = $true
  $openPageButton = New-Object System.Windows.Forms.Button
  $openPageButton.Text = "Open Download Page"
  $openPageButton.AutoSize = $true
  $reloadPackagesButton = New-Object System.Windows.Forms.Button
  $reloadPackagesButton.Text = "Reload List"
  $reloadPackagesButton.AutoSize = $true

  $installButtonsPanel.Controls.Add($installButton)
  $installButtonsPanel.Controls.Add($openPageButton)
  $installButtonsPanel.Controls.Add($reloadPackagesButton)

  $installStatus = New-Object System.Windows.Forms.TextBox
  $installStatus.Multiline = $true
  $installStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
  $installStatus.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
  $installStatus.ReadOnly = $true
  $installStatus.BackColor = [System.Drawing.Color]::White

  $installBottom.Controls.Add($installButtonsPanel,0,0)
  $installBottom.Controls.Add($installStatus,0,1)

  $installLayout.Controls.Add($packageList,0,0)
  $installLayout.Controls.Add($installBottom,0,1)
  $installPanel.Controls.Add($installLayout)

  function Populate-PackageList {
      param([switch]$ForceRefresh)

      $packages = Get-InstallPackages @PSBoundParameters

      $packageList.BeginUpdate()
      try {
          $packageList.Items.Clear()
          foreach ($pkg in $packages) {
              $item = New-Object System.Windows.Forms.ListViewItem($pkg.Name)
              [void]$item.SubItems.Add($pkg.Description)
              [void]$item.SubItems.Add($pkg.Url)
              [void]$item.SubItems.Add($pkg.Arguments)
              $item.Tag = $pkg
              [void]$packageList.Items.Add($item)
          }
      } finally {
          $packageList.EndUpdate()
      }

      if ($packageList.Items.Count -gt 0) {
          $packageList.Items[0].Selected = $true
          $packageList.Items[0].EnsureVisible()
      }

      $shouldAnnounce = $ForceRefresh -or -not $script:InstallPackagesHasAnnounced
      if ($shouldAnnounce) {
          $message = $null
          if ($script:InstallPackagesSource -eq 'Remote') {
              $host = $null
              try {
                  $uri = [System.Uri]$script:PackageService.Endpoint
                  $host = $uri.Host
              } catch {
                  $host = $script:PackageService.Endpoint
              }
              $message = "Loaded $($packages.Count) package(s) from Google Apps Script ($host)."
          } elseif ($script:InstallPackagesLastError) {
              $errorText = $script:InstallPackagesLastError
              if ($errorText.Length -gt 140) {
                  $errorText = $errorText.Substring(0,140) + "..."
              }
              $message = "Using built-in package catalog because remote sync failed: $errorText"
          } elseif ($ForceRefresh) {
              $message = "Package list refreshed."
          } else {
              $message = "Using built-in package catalog."
          }

          if ($message -and $null -ne $installStatus) {
              Update-InstallStatus $message
              $script:InstallPackagesHasAnnounced = $true
          }
      }
  }

  function Update-InstallStatus {
      param([string]$Message)
      $timestamp = (Get-Date).ToString("HH:mm:ss")
      $installStatus.AppendText("[$timestamp] $Message`r`n")
      $installStatus.SelectionStart = $installStatus.Text.Length
      $installStatus.ScrollToCaret()
  }

  $installButton.Add_Click({
      if ($packageList.SelectedItems.Count -eq 0) {
          Update-InstallStatus "Select a package first."
          return
      }
      $pkg = $packageList.SelectedItems[0].Tag
      $sanitized = ($pkg.Name -replace '[^\w]+','_')
      if (-not $sanitized) { $sanitized = "package" }
      $extension = [System.IO.Path]::GetExtension($pkg.Url)
      if (-not $extension) { $extension = ".exe" }
      $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
      $tempFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("{0}-{1}{2}" -f $sanitized,$timestamp,$extension)
      $durationSeconds = $null
      try {
          Update-InstallStatus "Downloading $($pkg.Name)..."
          Download-File -Url $pkg.Url -Destination $tempFile
          Update-InstallStatus "Saved to $tempFile"
          Update-InstallStatus "Launching installer..."
          $startTime = Get-Date
          if ([string]::IsNullOrWhiteSpace($pkg.Arguments)) {
              Start-Process -FilePath $tempFile -Wait -ErrorAction Stop
          } else {
              Start-Process -FilePath $tempFile -ArgumentList $pkg.Arguments -Wait -ErrorAction Stop
          }
          $durationSeconds = ((Get-Date) - $startTime).TotalSeconds
          if ($durationSeconds -and $durationSeconds -gt 0) {
              Update-InstallStatus ("Installation finished for $($pkg.Name) in {0:N1} seconds." -f $durationSeconds)
          } else {
              Update-InstallStatus "Installation finished for $($pkg.Name)."
          }
          $successMessage = if ($durationSeconds -and $durationSeconds -gt 0) {
              ("Completed in {0:N1} seconds." -f $durationSeconds)
          } else {
              "Installer completed successfully."
          }
          Send-InstallReport -Package $pkg -Status "Success" -Message $successMessage -DurationSeconds $durationSeconds
      } catch {
          $errorMessage = $_.Exception.Message
          Update-InstallStatus "Failed: $errorMessage"
          Send-InstallReport -Package $pkg -Status "Failed" -Message $errorMessage -DurationSeconds $durationSeconds -ErrorDetail $_.Exception.ToString()
      }
  })

  $openPageButton.Add_Click({
      if ($packageList.SelectedItems.Count -eq 0) {
          Update-InstallStatus "Select a package to open."
          return
      }
      $pkg = $packageList.SelectedItems[0].Tag
      try {
          Start-Process $pkg.Url
          Update-InstallStatus "Opened $($pkg.Url) in default browser."
      } catch {
          Update-InstallStatus "Unable to open URL: $($_.Exception.Message)"
      }
  })

  $reloadPackagesButton.Add_Click({ Populate-PackageList -ForceRefresh })

  $searchPanel = New-Object System.Windows.Forms.Panel
  $searchPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $searchPanel.Visible = $false

  $searchLayout = New-Object System.Windows.Forms.TableLayoutPanel
  $searchLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $searchLayout.ColumnCount = 1
  $searchLayout.RowCount = 5
  $searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,45)))
  $searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,45)))
  $searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,45)))
  $searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
  $searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,50)))
  $searchLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))

  $rootPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $rootPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $rootPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $rootPanel.WrapContents = $false
  $rootPanel.Padding = New-Object System.Windows.Forms.Padding(5,10,5,0)

  $rootLabel = New-Object System.Windows.Forms.Label
  $rootLabel.Text = "Root folder:"
  $rootLabel.AutoSize = $true
  $rootText = New-Object System.Windows.Forms.TextBox
  $rootText.Width = 500
  $rootText.ReadOnly = $true
  $browseButton = New-Object System.Windows.Forms.Button
  $browseButton.Text = "Browse..."
  $browseButton.AutoSize = $true
  $rootPanel.Controls.Add($rootLabel)
  $rootPanel.Controls.Add($rootText)
  $rootPanel.Controls.Add($browseButton)

  $keywordPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $keywordPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $keywordPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $keywordPanel.WrapContents = $false
  $keywordPanel.Padding = New-Object System.Windows.Forms.Padding(5,5,5,0)

  $keywordLabel = New-Object System.Windows.Forms.Label
  $keywordLabel.Text = "Keyword:"
  $keywordLabel.AutoSize = $true
  $keywordBox = New-Object System.Windows.Forms.TextBox
  $keywordBox.Width = 260
  $modeLabel = New-Object System.Windows.Forms.Label
  $modeLabel.Text = "Mode:"
  $modeLabel.AutoSize = $true
  $modeCombo = New-Object System.Windows.Forms.ComboBox
  $modeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$modeCombo.Items.AddRange(@("Name","Content"))
  $modeCombo.SelectedIndex = 0
  $keywordPanel.Controls.Add($keywordLabel)
  $keywordPanel.Controls.Add($keywordBox)
  $keywordPanel.Controls.Add($modeLabel)
  $keywordPanel.Controls.Add($modeCombo)

  $searchButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $searchButtonsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $searchButtonsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $searchButtonsPanel.WrapContents = $false
  $searchButtonsPanel.Padding = New-Object System.Windows.Forms.Padding(5,5,5,0)

  $searchButton = New-Object System.Windows.Forms.Button
  $searchButton.Text = "Search"
  $searchButton.AutoSize = $true
  $cancelSearchButton = New-Object System.Windows.Forms.Button
  $cancelSearchButton.Text = "Cancel"
  $cancelSearchButton.AutoSize = $true
  $cancelSearchButton.Enabled = $false
  $searchButtonsPanel.Controls.Add($searchButton)
  $searchButtonsPanel.Controls.Add($cancelSearchButton)

  $searchResults = New-Object System.Windows.Forms.ListView
  $searchResults.Dock = [System.Windows.Forms.DockStyle]::Fill
  $searchResults.View = [System.Windows.Forms.View]::Details
  $searchResults.FullRowSelect = $true
  $searchResults.HideSelection = $false
  $searchResults.MultiSelect = $false
  [void]$searchResults.Columns.Add("Type",80)
  [void]$searchResults.Columns.Add("Path",520)
  [void]$searchResults.Columns.Add("Match",300)

  $searchStatus = New-Object System.Windows.Forms.Label
  $searchStatus.Text = "Ready."
  $searchStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
  $searchStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $searchStatus.Padding = New-Object System.Windows.Forms.Padding(5)

  $searchLayout.Controls.Add($rootPanel,0,0)
  $searchLayout.Controls.Add($keywordPanel,0,1)
  $searchLayout.Controls.Add($searchButtonsPanel,0,2)
  $searchLayout.Controls.Add($searchResults,0,3)
  $searchLayout.Controls.Add($searchStatus,0,4)
  $searchPanel.Controls.Add($searchLayout)

  $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
  $browseButton.Add_Click({
      if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
          $rootText.Text = $folderDialog.SelectedPath
      }
  })

  $searchResults.Add_DoubleClick({
      if ($searchResults.SelectedItems.Count -gt 0) {
          $path = $searchResults.SelectedItems[0].SubItems[1].Text
          if (Test-Path $path) {
              Start-Process explorer.exe "/select,`"$path`""
          }
      }
  })

  $cancelSearchButton.Add_Click({
      $script:SearchCancelled = $true
      $searchStatus.Text = "Cancelling..."
  })

  $searchButton.Add_Click({
      if (-not (Test-Path $rootText.Text)) {
          $searchStatus.Text = "Choose a valid root folder."
          return
      }
      if ([string]::IsNullOrWhiteSpace($keywordBox.Text)) {
          $searchStatus.Text = "Enter a keyword."
          return
      }
      $searchButton.Enabled = $false
      $cancelSearchButton.Enabled = $true
      $searchResults.Items.Clear()
      $searchStatus.Text = "Searching..."
      $script:SearchCancelled = $false
      try {
          $results = Invoke-FileSearch -Root $rootText.Text -Keyword $keywordBox.Text.Trim() -Mode $modeCombo.SelectedItem
          foreach ($result in $results) {
              if ($script:SearchCancelled) { break }
              $item = New-Object System.Windows.Forms.ListViewItem($result.Type)
              [void]$item.SubItems.Add($result.Path)
              [void]$item.SubItems.Add($result.Match)
              $item.Tag = $result
              [void]$searchResults.Items.Add($item)
          }
          if ($script:SearchCancelled) {
              $searchStatus.Text = "Search cancelled."
          } else {
              $searchStatus.Text = "Search finished. Matches: $($searchResults.Items.Count)"
          }
      } catch {
          $searchStatus.Text = "Search failed: $($_.Exception.Message)"
      } finally {
          $searchButton.Enabled = $true
          $cancelSearchButton.Enabled = $false
          $script:SearchCancelled = $false
      }
  })

  $startupPanel = New-Object System.Windows.Forms.Panel
  $startupPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $startupPanel.Visible = $false

  $startupLayout = New-Object System.Windows.Forms.TableLayoutPanel
  $startupLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $startupLayout.ColumnCount = 1
  $startupLayout.RowCount = 3
  $startupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,65)))
  $startupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,130)))
  $startupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
  $startupLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))

  $startupList = New-Object System.Windows.Forms.ListView
  $startupList.View = [System.Windows.Forms.View]::Details
  $startupList.FullRowSelect = $true
  $startupList.HideSelection = $false
  $startupList.MultiSelect = $false
  $startupList.Dock = [System.Windows.Forms.DockStyle]::Fill
  [void]$startupList.Columns.Add("Name",170)
  [void]$startupList.Columns.Add("Command",460)
  [void]$startupList.Columns.Add("Scope",120)

  $startupDetails = New-Object System.Windows.Forms.TableLayoutPanel
  $startupDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
  $startupDetails.ColumnCount = 4
  $startupDetails.RowCount = 3
  $startupDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,18)))
  $startupDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,32)))
  $startupDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,18)))
  $startupDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,32)))
  $startupDetails.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,35)))
  $startupDetails.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,35)))
  $startupDetails.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,60)))

  $nameLabel = New-Object System.Windows.Forms.Label
  $nameLabel.Text = "Program name:"
  $nameLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $nameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $nameBox = New-Object System.Windows.Forms.TextBox
  $nameBox.Dock = [System.Windows.Forms.DockStyle]::Fill

  $startupDetails.Controls.Add($nameLabel,0,0)
  $startupDetails.Controls.Add($nameBox,1,0)
  $startupDetails.SetColumnSpan($nameBox,3)

  $pathLabel = New-Object System.Windows.Forms.Label
  $pathLabel.Text = "Executable path:"
  $pathLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $pathLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $pathBox = New-Object System.Windows.Forms.TextBox
  $pathBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $pathBrowse = New-Object System.Windows.Forms.Button
  $pathBrowse.Text = "Browse..."
  $pathBrowse.Dock = [System.Windows.Forms.DockStyle]::Fill

  $startupDetails.Controls.Add($pathLabel,0,1)
  $startupDetails.Controls.Add($pathBox,1,1)
  $startupDetails.SetColumnSpan($pathBox,2)
  $startupDetails.Controls.Add($pathBrowse,3,1)

  $scopeLabel = New-Object System.Windows.Forms.Label
  $scopeLabel.Text = "Scope:"
  $scopeLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $scopeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $scopeCombo = New-Object System.Windows.Forms.ComboBox
  $scopeCombo.Dock = [System.Windows.Forms.DockStyle]::Fill
  $scopeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$scopeCombo.Items.Add("CurrentUser")
  [void]$scopeCombo.Items.Add("LocalMachine")
  $scopeCombo.SelectedIndex = 0
  $addStartupButton = New-Object System.Windows.Forms.Button
  $addStartupButton.Text = "Add / Update"
  $addStartupButton.Dock = [System.Windows.Forms.DockStyle]::Fill
  $actionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $actionsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $actionsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $actionsPanel.WrapContents = $false
  $actionsPanel.Padding = New-Object System.Windows.Forms.Padding(0)

  $removeStartupButton = New-Object System.Windows.Forms.Button
  $removeStartupButton.Text = "Remove Selected"
  $removeStartupButton.AutoSize = $true
  $refreshStartupButton = New-Object System.Windows.Forms.Button
  $refreshStartupButton.Text = "Refresh"
  $refreshStartupButton.AutoSize = $true
  $actionsPanel.Controls.Add($removeStartupButton)
  $actionsPanel.Controls.Add($refreshStartupButton)

  $startupDetails.Controls.Add($scopeLabel,0,2)
  $startupDetails.Controls.Add($scopeCombo,1,2)
  $startupDetails.Controls.Add($addStartupButton,2,2)
  $startupDetails.Controls.Add($actionsPanel,3,2)

  $startupStatus = New-Object System.Windows.Forms.Label
  $startupStatus.Text = "Ready."
  $startupStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
  $startupStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $startupStatus.Padding = New-Object System.Windows.Forms.Padding(5)

  $startupLayout.Controls.Add($startupList,0,0)
  $startupLayout.Controls.Add($startupDetails,0,1)
  $startupLayout.Controls.Add($startupStatus,0,2)
  $startupPanel.Controls.Add($startupLayout)

  $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $fileDialog.Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*"
  $pathBrowse.Add_Click({
      if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
          $pathBox.Text = $fileDialog.FileName
      }
  })

  function Refresh-StartupList {
      $startupList.Items.Clear()
      $entries = Get-StartupEntries
      foreach ($entry in $entries) {
          $item = New-Object System.Windows.Forms.ListViewItem($entry.Name)
          [void]$item.SubItems.Add($entry.Command)
          [void]$item.SubItems.Add($entry.Scope)
          $item.Tag = $entry
          [void]$startupList.Items.Add($item)
      }
      $startupStatus.Text = "Entries loaded: $($startupList.Items.Count)"
  }

  $startupList.Add_SelectedIndexChanged({
      if ($startupList.SelectedItems.Count -gt 0) {
          $entry = $startupList.SelectedItems[0].Tag
          $nameBox.Text = $entry.Name
          $pathBox.Text = $entry.Command
          $scopeCombo.SelectedItem = $entry.Scope
      }
  })

  $addStartupButton.Add_Click({
      if ([string]::IsNullOrWhiteSpace($nameBox.Text)) {
          $startupStatus.Text = "Enter a program name."
          return
      }
      if ([string]::IsNullOrWhiteSpace($pathBox.Text)) {
          $startupStatus.Text = "Enter or browse to the executable."
          return
      }
      $scope = $scopeCombo.SelectedItem
      try {
          Add-StartupEntry -Name $nameBox.Text.Trim() -Command $pathBox.Text.Trim() -Scope $scope
          $startupStatus.Text = "Startup entry saved for $scope."
          Refresh-StartupList
      } catch {
          $startupStatus.Text = "Failed to save entry: $($_.Exception.Message)"
      }
  })

  $removeStartupButton.Add_Click({
      if ($startupList.SelectedItems.Count -eq 0) {
          $startupStatus.Text = "Select an entry to remove."
          return
      }
      $entry = $startupList.SelectedItems[0].Tag
      try {
          Remove-StartupEntry -Name $entry.Name -RegistryPath $entry.RegistryPath
          $startupStatus.Text = "Removed entry '$($entry.Name)'."
          Refresh-StartupList
      } catch {
          $startupStatus.Text = "Failed to remove entry: $($_.Exception.Message)"
      }
  })

  $refreshStartupButton.Add_Click({ Refresh-StartupList })

  $script:ContentPanels = @{
      System = $systemPanel
      Install = $installPanel
      Search  = $searchPanel
      Startup = $startupPanel
  }
  foreach ($panel in $script:ContentPanels.GetEnumerator()) {
      $panel.Value.Visible = $false
      $contentPanel.Controls.Add($panel.Value)
  }

  function Highlight-NavButton {
      param([string]$ActiveKey)
      foreach ($btn in $buttons) {
          if ($btn.Tag -eq $ActiveKey) {
              $btn.BackColor = [System.Drawing.Color]::FromArgb(200,220,255)
          } else {
              $btn.BackColor = [System.Drawing.Color]::White
          }
      }
  }

  function Show-Panel {
      param([string]$Key)
      foreach ($entry in $script:ContentPanels.GetEnumerator()) {
          $entry.Value.Visible = $false
      }
      if ($script:ContentPanels.ContainsKey($Key)) {
          $panel = $script:ContentPanels[$Key]
          $panel.Visible = $true
          $panel.BringToFront()
          Highlight-NavButton $Key
      }
  }

foreach ($button in $buttons) {
    # 각 버튼의 Tag 값은 "System" / "Install" / "Search" / "Startup" 이어야 함
    $button.Add_Click({
        param($sender, $e)

        Show-Panel -Key $sender.Tag

        switch ($sender.Tag) {
            "System"  { Update-HardwareView }
            "Install" { Populate-PackageList }
            "Search"  { $searchStatus.Text = "Ready. Select a folder and enter a keyword." }
            "Startup" { Refresh-StartupList }
            default   { }  # 필요시 기본 동작
        }
    })
}

  Show-Panel "System"
  Update-HardwareView
  Populate-PackageList

  [System.Windows.Forms.Application]::Run($form)
