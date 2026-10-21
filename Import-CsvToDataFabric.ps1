<#
.SYNOPSIS
    Creates a UiPath Data Fabric entity from an arbitrary CSV file and imports its rows.

.DESCRIPTION
    1. Parses the CSV (quoted values, embedded newlines, ; , TAB | delimiters, duplicate/empty headers).
    2. Turns every header into a valid Data Fabric field name
       (letters+digits only, starts with a letter, 3-100 chars, no reserved/system/keyword names, unique).
       The original header is kept as the field's display name, so the entity still reads naturally.
    3. Infers a field type per column (BOOLEAN, DECIMAL, DATE, DATETIME_WITH_TZ, STRING, MULTILINE_*)
       from ALL rows. Field types cannot be changed after creation, so the plan is shown for confirmation.
    4. Creates the entity with `uip df entities create`, reads the schema back with `uip df entities get`,
       rewrites the CSV with the server's display names as headers and normalized values, and loads it
       with `uip df records import` (in chunks).

    Requires the uip CLI, logged in (`uip login`), with the Data Fabric tool
    (`uip tools install @uipath/data-fabric-tool@latest`). Works in Windows PowerShell 5.1 and PowerShell 7.

.PARAMETER CsvPath
    CSV file to import.
.PARAMETER EntityName
    Entity name. Default: derived from the file name. Letters, digits, underscores.
.PARAMETER EntityDisplayName
    Entity display name. Default: the file name.
.PARAMETER Description
    Entity description.
.PARAMETER FolderKey
    Orchestrator folder key (GUID) for a folder-scoped entity. Omit for tenant level.
    Find it with: uip or folders list --output json
.PARAMETER Delimiter
    Column delimiter. Default: auto-detected from the header line (, ; TAB |).
.PARAMETER Encoding
    File encoding when the file has no BOM. Default UTF8. Use 1252 for classic German Excel exports.
.PARAMETER DecimalComma
    Numbers use a decimal comma (1.234,56). Values are converted to 1234.56 on import.
.PARAMETER DateFormats
    Accepted date-only formats. Values are converted to yyyy-MM-dd on import.
.PARAMETER ColumnTypes
    Overrides the inferred type per column, keyed by original header or field name.
    Values: STRING, MULTILINE_TEXT, MULTILINE_MAX, BOOLEAN, DATE, DATETIME_WITH_TZ, DECIMAL or DECIMAL:<precision>.
    Example: @{ 'PLZ' = 'STRING'; 'Betrag' = 'DECIMAL:2' }
.PARAMETER AllStrings
    Skip type inference; every column becomes STRING (or MULTILINE_* for long values).
.PARAMETER PlainDisplayNames
    Use the sanitized field name as display name instead of the original header.
.PARAMETER UseExisting
    Import into an existing entity with the same name instead of creating one. Columns are matched
    by field name or display name; unmatched columns are skipped. Stops with an error if the entity
    does not exist in the given scope (tenant, or -FolderKey).
.PARAMETER ChunkSize
    Rows per `records import` call. 0 = one call.
.PARAMETER Force
    Do not ask for confirmation before creating the entity.
.PARAMETER DryRun
    Only show the plan and write the cleaned CSV + entity JSON; change nothing in the tenant.
.PARAMETER SchemaOnly
    Create the entity from the CSV's columns and types, but import no rows.
.PARAMETER KeepTemp
    Keep the generated JSON/CSV files (their folder is printed).

.EXAMPLE
    .\Import-CsvToDataFabric.ps1 -CsvPath .\customers.csv -DryRun
.EXAMPLE
    .\Import-CsvToDataFabric.ps1 -CsvPath .\customers.csv -SchemaOnly
.EXAMPLE
    .\Import-CsvToDataFabric.ps1 -CsvPath .\Kunden.csv -Encoding 1252 -DecimalComma -ColumnTypes @{ PLZ = 'STRING' }
.EXAMPLE
    .\Import-CsvToDataFabric.ps1 -CsvPath .\orders.csv -EntityName Orders_2026 -FolderKey 3f2a...-... -Force
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$CsvPath,
    [string]$EntityName,
    [string]$EntityDisplayName,
    [string]$Description = '',
    [string]$FolderKey,
    [string]$Delimiter,
    [string]$Encoding = 'UTF8',
    [switch]$DecimalComma,
    [string[]]$DateFormats = @('yyyy-MM-dd', 'dd.MM.yyyy'),
    [hashtable]$ColumnTypes = @{},
    [switch]$AllStrings,
    [switch]$PlainDisplayNames,
    [switch]$UseExisting,
    [int]$ChunkSize = 5000,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$SchemaOnly,
    [switch]$KeepTemp
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture

# --- Name rules -------------------------------------------------------------------------------

$SystemFields = @('Id', 'CreatedBy', 'CreateTime', 'UpdatedBy', 'UpdateTime', 'RecordOwner')
$SystemDisplayNames = @('Id', 'Created By', 'Create Time', 'Updated By', 'Update Time', 'Record Owner') + $SystemFields

# C# + VB.NET keywords and SQL words the service rejects. The server stays the final judge;
# renaming a harmless name only changes the internal field name, the display name keeps the header.
$Keywords = @(
    'abstract', 'as', 'base', 'bool', 'break', 'byte', 'case', 'catch', 'char', 'checked', 'class', 'const',
    'continue', 'decimal', 'default', 'delegate', 'do', 'double', 'else', 'enum', 'event', 'explicit', 'extern',
    'false', 'finally', 'fixed', 'float', 'for', 'foreach', 'goto', 'if', 'implicit', 'in', 'int', 'interface',
    'internal', 'is', 'lock', 'long', 'namespace', 'new', 'null', 'object', 'operator', 'out', 'override',
    'params', 'private', 'protected', 'public', 'readonly', 'ref', 'return', 'sbyte', 'sealed', 'short',
    'sizeof', 'stackalloc', 'static', 'string', 'struct', 'switch', 'this', 'throw', 'true', 'try', 'typeof',
    'uint', 'ulong', 'unchecked', 'unsafe', 'ushort', 'using', 'virtual', 'void', 'volatile', 'while',
    'addhandler', 'addressof', 'alias', 'and', 'andalso', 'boolean', 'byref', 'byval', 'call', 'cbool', 'cbyte',
    'cchar', 'cdate', 'cdbl', 'cdec', 'cint', 'clng', 'cobj', 'csbyte', 'cshort', 'csng', 'cstr', 'ctype',
    'cuint', 'culng', 'cushort', 'date', 'declare', 'dim', 'directcast', 'each', 'elseif', 'end', 'endif',
    'erase', 'error', 'exit', 'friend', 'function', 'get', 'gettype', 'getxmlnamespace', 'global', 'gosub',
    'handles', 'implements', 'imports', 'inherits', 'integer', 'isnot', 'let', 'lib', 'like', 'loop', 'me',
    'mod', 'module', 'mustinherit', 'mustoverride', 'mybase', 'myclass', 'narrowing', 'next', 'not', 'nothing',
    'notinheritable', 'notoverridable', 'of', 'on', 'option', 'optional', 'or', 'orelse', 'overloads',
    'overridable', 'overrides', 'paramarray', 'partial', 'property', 'raiseevent', 'redim', 'rem',
    'removehandler', 'resume', 'select', 'set', 'shadows', 'shared', 'single', 'step', 'stop', 'sub',
    'synclock', 'then', 'to', 'trycast', 'variant', 'wend', 'when', 'widening', 'with', 'withevents',
    'writeonly', 'xor', 'order', 'group', 'table', 'from', 'where', 'by', 'index', 'view'
)

function ConvertTo-Ascii([string]$s) {
    # Char codes keep this file ASCII (Windows PowerShell 5.1 reads BOM-less scripts as ANSI).
    $map = @(
        @(0x00E4, 'ae'), @(0x00F6, 'oe'), @(0x00FC, 'ue'), @(0x00C4, 'Ae'), @(0x00D6, 'Oe'), @(0x00DC, 'Ue'),
        @(0x00DF, 'ss'), @(0x20AC, 'Eur'), @(0x0025, 'Pct'), @(0x0023, 'No')
    )
    foreach ($p in $map) { $s = $s.Replace([string][char]$p[0], $p[1]) }
    $d = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $d.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    $sb.ToString()
}

# Header -> PascalCase letters+digits. $allowUnderscore for entity names.
function ConvertTo-SafeName([string]$raw, [bool]$allowUnderscore, [string]$fallback) {
    $s = ConvertTo-Ascii $raw
    if ($allowUnderscore) {
        $s = ($s -replace '[^A-Za-z0-9_]+', '_').Trim('_')
    } else {
        $parts = $s -split '[^A-Za-z0-9]+' | Where-Object { $_ -ne '' }
        # ALL-CAPS tokens (SAP-style COMPANY_CODE) become CompanyCode; the server would otherwise store Companycode.
        $parts = $parts | ForEach-Object { if ($_ -cmatch '^[A-Z0-9]+$' -and $_.Length -gt 1) { $_.ToLowerInvariant() } else { $_ } }
        $s = ($parts | ForEach-Object { $_.Substring(0, 1).ToUpperInvariant() + $_.Substring(1) }) -join ''
    }
    if ($s -eq '') { $s = $fallback }
    if ($s -notmatch '^[A-Za-z]') { $s = 'F' + $s }
    if ($s.Length -lt 3) { $s = $s + 'Value' }
    if ($s.Length -gt 100) { $s = $s.Substring(0, 100) }
    $s
}

function Get-UniqueName([string]$name, [System.Collections.Generic.HashSet[string]]$taken, [int]$maxLen, [string]$separator = '') {
    $candidate = $name
    $i = 2
    while ($taken.Contains($candidate)) {
        $suffix = $separator + [string]$i
        $base = $name
        if ($base.Length + $suffix.Length -gt $maxLen) { $base = $base.Substring(0, $maxLen - $suffix.Length) }
        $candidate = $base + $suffix
        $i++
    }
    [void]$taken.Add($candidate)
    $candidate
}

# --- uip wrapper ------------------------------------------------------------------------------

function Invoke-Uip([string[]]$UipArgs) {
    $all = @($UipArgs) + @('--output', 'json')
    # uip writes notices ("Checking today's update...") to stderr. Windows PowerShell 5.1 turns those into
    # NativeCommandError records, which abort the script under ErrorActionPreference=Stop. Keep stderr apart.
    $prevPref = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $output = & uip @all 2>&1 } finally { $ErrorActionPreference = $prevPref }
    $code = $LASTEXITCODE
    $stdout = @($output | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" })
    $stderr = @($output | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" })
    $text = ($stdout -join "`n")
    if ($text.IndexOf('{') -lt 0) { $text = (@($stdout) + @($stderr)) -join "`n" }
    $start = $text.IndexOf('{'); $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -lt $start) { throw "uip $($UipArgs -join ' ') returned no JSON (exit $code):`n$text" }
    $json = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json
    if ($json.Result -ne 'Success') {
        throw "uip $($UipArgs[0..2] -join ' ') failed: $($text.Substring($start, $end - $start + 1))"
    }
    $json
}

function Add-FolderArgs([string[]]$a) {
    if ($FolderKey) { return @($a) + @('--folder-key', $FolderKey) }
    @($a)
}

# --- CSV reading ------------------------------------------------------------------------------

function Get-FileEncoding([string]$path, [string]$fallback) {
    $bytes = New-Object byte[] 4
    $fs = [IO.File]::OpenRead($path)
    try { [void]$fs.Read($bytes, 0, 4) } finally { $fs.Dispose() }
    if ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return New-Object Text.UTF8Encoding($false) }
    if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode }
    if ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode }
    $n = 0
    if ([int]::TryParse($fallback, [ref]$n)) { return [Text.Encoding]::GetEncoding($n) }
    if ($fallback -eq 'UTF8') { return New-Object Text.UTF8Encoding($false) }
    [Text.Encoding]::GetEncoding($fallback)
}

function Find-Delimiter([string]$headerLine) {
    $best = ','; $bestCount = -1
    foreach ($d in @(',', ';', "`t", '|')) {
        $inQuote = $false; $count = 0
        foreach ($ch in $headerLine.ToCharArray()) {
            if ($ch -eq '"') { $inQuote = -not $inQuote }
            elseif (-not $inQuote -and [string]$ch -eq $d) { $count++ }
        }
        if ($count -gt $bestCount) { $best = $d; $bestCount = $count }
    }
    $best
}

function Read-CsvRows([string]$path, [Text.Encoding]$enc, [string]$delim) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($path, $enc, $true)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters($delim)
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        $rows = New-Object System.Collections.Generic.List[string[]]
        while (-not $parser.EndOfData) {
            try { $f = $parser.ReadFields() }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] {
                throw "Malformed CSV at line $($parser.ErrorLineNumber): $($parser.ErrorLine)"
            }
            if ($null -eq $f) { continue }
            if ($f.Length -eq 1 -and $f[0] -eq '') { continue }   # blank line
            $rows.Add($f)
        }
        ,$rows
    } finally { $parser.Close() }
}

# --- Type inference ---------------------------------------------------------------------------

$MaxSafe = [decimal]9007199254740991
$DateTimePattern = '^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:?\d{2})?$'

function ConvertTo-InvariantNumber([string]$v) {
    $t = $v.Trim()
    if ($DecimalComma) {
        if ($t -notmatch '^[+-]?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?$') { return $null }
        $t = $t.Replace('.', '').Replace(',', '.')
    } else {
        if ($t -notmatch '^[+-]?\d+(\.\d+)?$') { return $null }
    }
    # Leading zeros (ZIP codes, article numbers) must stay text.
    if ($t -match '^[+-]?0\d') { return $null }
    $d = [decimal]0
    if (-not [decimal]::TryParse($t, [Globalization.NumberStyles]::Float, $Invariant, [ref]$d)) { return $null }
    if ([Math]::Abs($d) -gt $MaxSafe) { return $null }
    $t
}

function ConvertTo-IsoDate([string]$v) {
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($v.Trim(), [string[]]$DateFormats, $Invariant, [Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd', $Invariant)
    }
    $null
}

function ConvertTo-IsoDateTime([string]$v) {
    $t = $v.Trim()
    if ($t -notmatch $DateTimePattern) { return $null }
    $dto = [DateTimeOffset]::MinValue
    $styles = [Globalization.DateTimeStyles]::AssumeLocal
    if (-not [DateTimeOffset]::TryParse($t, $Invariant, $styles, [ref]$dto)) { return $null }
    $dto.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", $Invariant)
}

function Get-StringType([int]$maxLen) {
    if ($maxLen -le 4000) {
        $limit = [Math]::Min(4000, [Math]::Max(255, [int][Math]::Ceiling($maxLen * 1.5)))
        return @{ type = 'STRING'; lengthLimit = $limit }
    }
    if ($maxLen -le 10000) { return @{ type = 'MULTILINE_TEXT'; lengthLimit = 10000 } }
    @{ type = 'MULTILINE_MAX' }
}

function Get-ColumnType([string[]]$values) {
    $nonEmpty = @($values | Where-Object { $_ -ne $null -and $_.Trim() -ne '' })
    $maxLen = 0
    foreach ($v in $values) { if ($v -and $v.Length -gt $maxLen) { $maxLen = $v.Length } }
    if ($AllStrings -or $nonEmpty.Count -eq 0) { return Get-StringType $maxLen }

    if (@($nonEmpty | Where-Object { $_.Trim() -notmatch '^(?i:true|false)$' }).Count -eq 0) { return @{ type = 'BOOLEAN' } }

    $isNumber = $true; $precision = 0
    foreach ($v in $nonEmpty) {
        $n = ConvertTo-InvariantNumber $v
        if ($null -eq $n) { $isNumber = $false; break }
        $dot = $n.IndexOf('.')
        if ($dot -ge 0) { $precision = [Math]::Max($precision, $n.Length - $dot - 1) }
    }
    if ($isNumber -and $precision -le 10) { return @{ type = 'DECIMAL'; decimalPrecision = $precision } }

    if (@($nonEmpty | Where-Object { $null -eq (ConvertTo-IsoDate $_) }).Count -eq 0) { return @{ type = 'DATE' } }
    if (@($nonEmpty | Where-Object { $null -eq (ConvertTo-IsoDateTime $_) }).Count -eq 0) { return @{ type = 'DATETIME_WITH_TZ' } }

    Get-StringType $maxLen
}

function ConvertFrom-TypeOverride([string]$spec, [int]$maxLen) {
    $s = $spec.Trim().ToUpperInvariant()
    if ($s -match '^DECIMAL(:(\d+))?$') {
        $p = 0; if ($Matches[2]) { $p = [int]$Matches[2] }
        return @{ type = 'DECIMAL'; decimalPrecision = $p }
    }
    switch ($s) {
        'STRING' { return @{ type = 'STRING'; lengthLimit = [Math]::Min(4000, [Math]::Max(255, [int][Math]::Ceiling($maxLen * 1.5))) } }
        'MULTILINE_TEXT' { return @{ type = 'MULTILINE_TEXT'; lengthLimit = 10000 } }
        { $_ -in 'MULTILINE_MAX', 'BOOLEAN', 'DATE', 'DATETIME_WITH_TZ' } { return @{ type = $s } }
    }
    throw "Unsupported type override '$spec'. Use STRING, MULTILINE_TEXT, MULTILINE_MAX, BOOLEAN, DATE, DATETIME_WITH_TZ, DECIMAL[:n]."
}

function ConvertTo-ImportValue([string]$v, [string]$type) {
    if ($null -eq $v -or $v.Trim() -eq '') { return '' }
    switch ($type) {
        'BOOLEAN' { return $v.Trim().ToLowerInvariant() }
        'DECIMAL' { $n = ConvertTo-InvariantNumber $v; if ($null -ne $n) { return $n }; return $v }
        'DATE' { $d = ConvertTo-IsoDate $v; if ($d) { return $d }; return $v }
        'DATETIME_WITH_TZ' { $d = ConvertTo-IsoDateTime $v; if ($d) { return $d }; return $v }
    }
    $v
}

# Typed JSON value for records insert; $null = leave the field empty.
function ConvertTo-InsertValue([string]$v, [string]$type) {
    $s = ConvertTo-ImportValue $v $type
    if ($s -eq '') { return $null }
    if ($type -eq 'BOOLEAN' -and $s -match '^(true|false)$') { return ($s -eq 'true') }
    if ($type -eq 'DECIMAL' -and $null -ne (ConvertTo-InvariantNumber $s)) { return [decimal]::Parse($s, $Invariant) }
    $s
}

function ConvertTo-CsvLine([string[]]$fields) {
    ($fields | ForEach-Object { '"' + ([string]$_).Replace('"', '""') + '"' }) -join ','
}

# --- Main -------------------------------------------------------------------------------------

$CsvPath = (Resolve-Path -LiteralPath $CsvPath).Path
$enc = Get-FileEncoding $CsvPath $Encoding
if (-not $Delimiter) {
    $reader = New-Object IO.StreamReader($CsvPath, $enc, $true)
    try { $firstLine = $reader.ReadLine() } finally { $reader.Dispose() }
    $Delimiter = Find-Delimiter $firstLine
}
$rows = Read-CsvRows $CsvPath $enc $Delimiter
if ($rows.Count -lt 1) { throw "CSV '$CsvPath' is empty." }
$headers = $rows[0]
$dataRows = $rows.GetRange(1, $rows.Count - 1)
$delimName = $Delimiter; if ($Delimiter -eq "`t") { $delimName = 'TAB' }
Write-Host "Read $($dataRows.Count) data rows, $($headers.Length) columns (delimiter '$delimName', encoding $($enc.WebName))."

$badRows = @(); for ($r = 0; $r -lt $dataRows.Count; $r++) { if ($dataRows[$r].Length -ne $headers.Length) { $badRows += ($r + 2) } }
if ($badRows.Count -gt 0) {
    Write-Warning "$($badRows.Count) row(s) have a different column count than the header (lines: $(($badRows | Select-Object -First 10) -join ', ')). Missing cells are imported empty, extra cells are dropped."
}
if (@($headers | Where-Object { $_ -match '[\uFFFD]' }).Count -gt 0 -or @($dataRows | Select-Object -First 200 | Where-Object { ($_ -join '') -match '[\uFFFD]' }).Count -gt 0) {
    Write-Warning "The file contains undecodable characters. If it is an Excel export, retry with -Encoding 1252."
}

$fileBase = [IO.Path]::GetFileNameWithoutExtension($CsvPath)
if (-not $EntityName) { $EntityName = ConvertTo-SafeName $fileBase $true 'CsvImport' }
if ($EntityName -notmatch '^[A-Za-z][A-Za-z0-9_]{2,99}$') { throw "Entity name '$EntityName' is invalid: start with a letter, letters/digits/underscores, 3-100 chars." }
if (-not $EntityDisplayName) { $EntityDisplayName = $fileBase }

# Column plan
$takenNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$takenDisplay = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($s in $SystemFields) { [void]$takenNames.Add($s) }
foreach ($s in $SystemDisplayNames) { [void]$takenDisplay.Add($s) }
[void]$takenNames.Add($EntityName)

$columns = @()
for ($c = 0; $c -lt $headers.Length; $c++) {
    $header = ([string]$headers[$c]) -replace '^[\uFEFF]', ''
    $cleanHeader = ($header -replace '\s+', ' ').Trim()

    $name = ConvertTo-SafeName $cleanHeader $false ("Column" + ($c + 1))
    if ($SystemFields -contains $name) { $name = 'Source' + $name }
    elseif ($Keywords -contains $name.ToLowerInvariant() -or $name -ieq $EntityName) { $name = $name + 'Value' }
    $name = Get-UniqueName $name $takenNames 100

    $display = $cleanHeader
    if ($PlainDisplayNames -or $display -eq '') { $display = $name }
    if ($SystemDisplayNames -contains $display) { $display = 'Source ' + $display }
    if ($display.Length -gt 120) { $display = $display.Substring(0, 120) }
    $display = Get-UniqueName $display $takenDisplay 128 ' '

    $values = New-Object string[] $dataRows.Count
    for ($r = 0; $r -lt $dataRows.Count; $r++) { if ($c -lt $dataRows[$r].Length) { $values[$r] = $dataRows[$r][$c] } else { $values[$r] = '' } }
    $maxLen = 0; foreach ($v in $values) { if ($v -and $v.Length -gt $maxLen) { $maxLen = $v.Length } }

    $override = $null
    foreach ($k in @($header, $cleanHeader, $name)) { if ($k -and $ColumnTypes.ContainsKey($k)) { $override = $ColumnTypes[$k]; break } }
    if ($override) { $typeInfo = ConvertFrom-TypeOverride $override $maxLen; $source = 'override' }
    else { $typeInfo = Get-ColumnType $values; $source = 'inferred' }

    $sample = ($values | Where-Object { $_ -and $_.Trim() -ne '' } | Select-Object -First 1)
    if ($sample -and $sample.Length -gt 30) { $sample = $sample.Substring(0, 27) + '...' }

    $columns += [pscustomobject]@{
        Index = $c; Header = $header; Name = $name; DisplayName = $display
        Type = $typeInfo.type; TypeInfo = $typeInfo; Source = $source; Sample = $sample
        ImportHeader = $null; Skip = $false
    }
}

Write-Host ""
Write-Host "Entity: $EntityName  (display '$EntityDisplayName', scope: $(if ($FolderKey) { "folder $FolderKey" } else { 'tenant' }))"
$columns | Select-Object Header, Name, DisplayName,
    @{ n = 'Type'; e = { $t = $_.Type; if ($_.TypeInfo.ContainsKey('decimalPrecision')) { $t += ":$($_.TypeInfo.decimalPrecision)" } elseif ($_.TypeInfo.ContainsKey('lengthLimit')) { $t += "($($_.TypeInfo.lengthLimit))" }; $t } },
    Source, Sample | Format-Table -AutoSize | Out-Host

$tempDir = Join-Path ([IO.Path]::GetTempPath()) ("df-import-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempDir | Out-Null
$utf8 = New-Object Text.UTF8Encoding($false)

$fields = @($columns | ForEach-Object {
    $f = [ordered]@{ name = $_.Name; type = $_.Type; displayName = $_.DisplayName }
    foreach ($k in @('lengthLimit', 'decimalPrecision')) { if ($_.TypeInfo.ContainsKey($k)) { $f[$k] = $_.TypeInfo[$k] } }
    $f
})
$body = [ordered]@{ displayName = $EntityDisplayName; description = $Description; fields = $fields }
$bodyPath = Join-Path $tempDir 'entity.json'
[IO.File]::WriteAllText($bodyPath, ($body | ConvertTo-Json -Depth 6), $utf8)

# records import ends a row at every line break, even inside quotes. Rows with a multi-line value
# therefore go through records insert (JSON) instead; all other rows use the bulk CSV import.
$multiRows = New-Object System.Collections.Generic.List[string[]]
$bulkRows = New-Object System.Collections.Generic.List[string[]]
foreach ($row in $dataRows) {
    if (@($row | Where-Object { $_ -match '[\r\n]' }).Count -gt 0) { $multiRows.Add($row) } else { $bulkRows.Add($row) }
}
if ($multiRows.Count -gt 0) { Write-Host "$($multiRows.Count) row(s) contain line breaks inside a value; they are inserted with records insert." }

function Write-ImportFiles([object[]]$cols, $rowsToWrite) {
    $active = @($cols | Where-Object { -not $_.Skip })
    $headerLine = ConvertTo-CsvLine ([string[]]@($active | ForEach-Object { $_.ImportHeader }))
    $size = $ChunkSize; if ($size -le 0) { $size = [Math]::Max(1, $rowsToWrite.Count) }
    $files = @()
    for ($start = 0; $start -lt $rowsToWrite.Count; $start += $size) {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add($headerLine)
        $stop = [Math]::Min($rowsToWrite.Count, $start + $size)
        for ($r = $start; $r -lt $stop; $r++) {
            $row = $rowsToWrite[$r]
            $cells = foreach ($col in $active) {
                $v = ''; if ($col.Index -lt $row.Length) { $v = $row[$col.Index] }
                ConvertTo-ImportValue $v $col.Type
            }
            $lines.Add((ConvertTo-CsvLine ([string[]]@($cells))))
        }
        $file = Join-Path $tempDir ("import-{0:D4}.csv" -f ($files.Count + 1))
        [IO.File]::WriteAllText($file, (($lines -join "`r`n") + "`r`n"), $utf8)
        $files += $file
    }
    $files
}

if ($DryRun) {
    foreach ($col in $columns) { $col.ImportHeader = $col.DisplayName }
    $files = @(Write-ImportFiles $columns $bulkRows)
    Write-Host "Dry run - nothing changed in the tenant."
    Write-Host "Entity definition: $bodyPath"
    Write-Host "Cleaned CSV:       $(if ($files.Count) { $files -join ', ' } else { '(none)' })"
    if ($multiRows.Count -gt 0) { Write-Host "Multi-line rows:   $($multiRows.Count) (sent via records insert)" }
    return
}

# Pre-flight: login + existing entity
$login = Invoke-Uip @('login', 'status')
Write-Host "Tenant: $($login.Data.Organization)/$($login.Data.Tenant) as $($login.Data.UserName)"

$existing = @((Invoke-Uip (Add-FolderArgs @('df', 'entities', 'list', '--native-only'))).Data | Where-Object { $_.Name -ieq $EntityName })
$entityId = $null
$created = $false
if ($existing.Count -gt 0) {
    if (-not $UseExisting) { throw "Entity '$EntityName' already exists (Id $($existing[0].Id)). Pass -UseExisting to import into it, or -EntityName for a new one." }
    $entityId = $existing[0].Id
    Write-Host "Using existing entity $EntityName ($entityId)."
} else {
    if ($UseExisting) {
        $scope = 'at tenant level'; if ($FolderKey) { $scope = "in folder $FolderKey" }
        throw "Entity '$EntityName' does not exist $scope. -UseExisting never creates an entity; check -EntityName / -FolderKey, or drop -UseExisting to create it."
    }
    if (-not $Force) {
        Write-Host "Field types cannot be changed after the entity is created."
        $what = "and import $($dataRows.Count) rows"; if ($SchemaOnly) { $what = "(schema only, no rows)" }
        $answer = Read-Host "Create entity '$EntityName' with $($columns.Count) fields $what? [y/N]"
        if ($answer -notmatch '^(y|yes|j|ja)$') { Write-Host "Aborted. Entity definition left at $bodyPath"; return }
    }
    $res = Invoke-Uip (Add-FolderArgs @('df', 'entities', 'create', $EntityName, '--file', $bodyPath))
    $entityId = $res.Data.Id
    $created = $true
    Write-Host "Created entity $EntityName ($entityId)."
}

if ($SchemaOnly) {
    if (-not $created) { Write-Host "Schema only: entity already existed, nothing changed." }
    if ($KeepTemp) { Write-Host "Generated files kept in $tempDir" } else { Remove-Item -LiteralPath $tempDir -Recurse -Force }
    return [pscustomobject]@{ EntityName = $EntityName; EntityId = $entityId; Created = $created; FolderKey = $FolderKey; RowsInCsv = $dataRows.Count; InsertedRecords = 0 }
}

# Read the schema back: the import header must be the server's DisplayName for each field.
$schema = (Invoke-Uip (Add-FolderArgs @('df', 'entities', 'get', $entityId))).Data
$serverFields = @($schema.Fields | Where-Object { $SystemFields -notcontains $_.Name })
$complex = @('CHOICE_SET_SINGLE', 'CHOICE_SET_MULTIPLE', 'RELATIONSHIP', 'FILE', 'AUTO_NUMBER')
foreach ($col in $columns) {
    $sf = $serverFields | Where-Object { $_.Name -ieq $col.Name } | Select-Object -First 1
    if (-not $sf) { $sf = $serverFields | Where-Object { $_.DisplayName -ieq $col.DisplayName -or $_.DisplayName -ieq $col.Header } | Select-Object -First 1 }
    if (-not $sf) { Write-Warning "Column '$($col.Header)' has no matching field in the entity - skipped."; $col.Skip = $true; continue }
    if ($complex -contains $sf.FieldDataType.Name) { Write-Warning "Field '$($sf.Name)' is $($sf.FieldDataType.Name); records import ignores this type - skipped."; $col.Skip = $true; continue }
    $col.ImportHeader = $sf.DisplayName
    $col | Add-Member -NotePropertyName FieldName -NotePropertyValue $sf.Name -Force
    $col.Type = $sf.FieldDataType.Name
}
$mappedNames = @($columns | Where-Object { -not $_.Skip } | ForEach-Object { $_.ImportHeader })
foreach ($sf in $serverFields) {
    if ($sf.IsRequired -and $mappedNames -notcontains $sf.DisplayName -and -not $sf.DefaultValue) { Write-Warning "Required field '$($sf.Name)' is not in the CSV - every row will fail." }
}
if (@($columns | Where-Object { -not $_.Skip }).Count -eq 0) { throw "No CSV column maps to a field of entity $EntityName." }

$files = @(Write-ImportFiles $columns $bulkRows)
$inserted = 0; $total = 0; $errorLinks = @(); $failedChunks = @(); $insertFailures = @()
$i = 0
foreach ($file in $files) {
    $i++
    Write-Host "Importing chunk $i/$($files.Count)..."
    try {
        $imp = Invoke-Uip (Add-FolderArgs @('df', 'records', 'import', $entityId, '--file', $file))
        $inserted += [int]$imp.Data.InsertedRecords
        $total += [int]$imp.Data.TotalRecords
        if ($imp.Data.PSObject.Properties['ErrorFileLink'] -and $imp.Data.ErrorFileLink) { $errorLinks += $imp.Data.ErrorFileLink }
    } catch {
        Write-Warning "Chunk $i failed: $($_.Exception.Message)"
        $failedChunks += $file
    }
}

$insertBatch = 100
$activeCols = @($columns | Where-Object { -not $_.Skip })
for ($start = 0; $start -lt $multiRows.Count; $start += $insertBatch) {
    $stop = [Math]::Min($multiRows.Count, $start + $insertBatch)
    $records = @()
    for ($r = $start; $r -lt $stop; $r++) {
        $row = $multiRows[$r]
        $rec = [ordered]@{}
        foreach ($col in $activeCols) {
            $v = ''; if ($col.Index -lt $row.Length) { $v = $row[$col.Index] }
            $val = ConvertTo-InsertValue $v $col.Type
            if ($null -ne $val) { $rec[$col.FieldName] = $val }
        }
        $records += $rec
    }
    $file = Join-Path $tempDir ("insert-{0:D4}.json" -f ([int]($start / $insertBatch) + 1))
    [IO.File]::WriteAllText($file, (ConvertTo-Json -InputObject @($records) -Depth 3 -Compress), $utf8)
    Write-Host "Inserting multi-line rows $($start + 1)-$stop..."
    $total += $records.Count
    try {
        $ins = Invoke-Uip (Add-FolderArgs @('df', 'records', 'insert', $entityId, '--file', $file))
        if ($records.Count -eq 1) { $inserted += 1 }
        else {
            $inserted += [int]$ins.Data.SuccessCount
            if ([int]$ins.Data.FailureCount -gt 0) { $insertFailures += @($ins.Data.FailureRecords) }
        }
    } catch {
        Write-Warning "Insert batch starting at multi-line row $($start + 1) failed: $($_.Exception.Message)"
        $failedChunks += $file
    }
}
$insertFailureFile = $null
if ($insertFailures.Count -gt 0) {
    $insertFailureFile = Join-Path $tempDir 'insert-failures.json'
    [IO.File]::WriteAllText($insertFailureFile, (ConvertTo-Json -InputObject @($insertFailures) -Depth 6), $utf8)
    Write-Warning "$($insertFailures.Count) multi-line row(s) were rejected; details in $insertFailureFile"
}

$result = [pscustomobject]@{
    EntityName      = $EntityName
    EntityId        = $entityId
    Created         = $created
    FolderKey       = $FolderKey
    RowsInCsv       = $dataRows.Count
    InsertedRecords = $inserted
    TotalRecords    = $total
    ErrorFileLinks  = $errorLinks
    InsertFailureFile = $insertFailureFile
    FailedChunks    = $failedChunks
}
if ($failedChunks.Count -gt 0 -or $errorLinks.Count -gt 0 -or $insertFailures.Count -gt 0 -or $inserted -lt $dataRows.Count) {
    Write-Warning "Not every row was imported. Download the ErrorFileLinks for row errors. Re-run a failed CSV chunk with: uip df records import $entityId --file <chunk>$(if ($FolderKey) { " --folder-key $FolderKey" }) --output json (insert-*.json files: uip df records insert ... --file <file>)"
    $KeepTemp = $true
}
if ($KeepTemp) { Write-Host "Generated files kept in $tempDir" } else { Remove-Item -LiteralPath $tempDir -Recurse -Force }
$result
