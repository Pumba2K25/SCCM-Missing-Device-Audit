<#
.SYNOPSIS
    Audits a list of Windows devices against Microsoft Configuration Manager (SCCM/MECM),
    Active Directory, and a single ICMP ping, then creates CSV and Excel reports.

.DESCRIPTION
    Designed for physical device-reconciliation and missing-device audits.

    The script:
      - Reads device names from a CSV
      - Matches devices in SCCM / Configuration Manager
      - Matches computer objects in Active Directory
      - Resolves the best available user/owner and manager
      - Sends exactly one sequential ping per device
      - Produces a technical CSV
      - Produces an Excel workbook with Dashboard, By Manager, Online, Offline,
        Retirement Review, and Manager Summary worksheets

    SCCM and Active Directory operations are read-only.

.NOTES
    PowerShell 5.1
    Requires:
      - ActiveDirectory PowerShell module
      - ConfigurationManager PowerShell module / ConfigMgr console
      - Microsoft Excel desktop application for XLSX generation

    Review the script and test it in your own environment before production use.
#>

[CmdletBinding()]
param(
    # Optional.
    # If blank, the newest usable CSV in the script folder is used.
    [Parameter(Mandatory = $false)]
    [string]$InputCsv,

    # Optional.
    # If blank, an XLSX report is created in the script folder.
    [Parameter(Mandatory = $false)]
    [string]$OutputWorkbook,

    # Optional if your SCCM CMSite drive already exists.
    [Parameter(Mandatory = $false)]
    [string]$SiteCode,

    [Parameter(Mandatory = $false)]
    [string]$SiteServer,

    # ONE ping per device.
    [Parameter(Mandatory = $false)]
    [int]$PingTimeoutMs = 750
)

# ============================================================
# LOST DEVICE AUDIT
# CSV INPUT -> AD + SCCM + ONE PING -> XLSX + RAW CSV
#
# PowerShell 5.1
#
# INPUT:
#   CSV file
#
# OUTPUT:
#   Lost_Device_Audit_YYYYMMDD_HHmm.xlsx
#   Lost_Device_Audit_YYYYMMDD_HHmm.csv
#
# NETWORK:
#   EXACTLY ONE sequential ping per device.
#   NO parallel pinging.
#   NO retries.
#
# AD / SCCM:
#   READ ONLY
# ============================================================

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'


# ============================================================
# GENERAL HELPERS
# ============================================================

function Get-SafePropertyValue {
    param(
        $Object,
        [string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]

    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}


function Convert-ToDisplayText {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [System.Array]) {
        return (($Value | Where-Object { $_ }) -join '; ')
    }

    return [string]$Value
}


function Escape-LdapValue {
    param(
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    # PowerShell 5.1-safe string replacements.
    $escaped = [string]$Value

    $escaped = $escaped.Replace(
        [string]'\',
        [string]'\5c'
    )

    $escaped = $escaped.Replace(
        [string]'*',
        [string]'\2a'
    )

    $escaped = $escaped.Replace(
        [string]'(',
        [string]'\28'
    )

    $escaped = $escaped.Replace(
        [string]')',
        [string]'\29'
    )

    $escaped = $escaped.Replace(
        [string][char]0,
        [string]'\00'
    )

    return $escaped
}


# ============================================================
# CSV INPUT
# ============================================================

function Get-LatestInputCsv {
    param(
        [string]$Folder
    )

    Write-Host ""
    Write-Host "Looking for newest input CSV..." -ForegroundColor Cyan

    $files = @(
        Get-ChildItem `
            -Path $Folder `
            -File `
            -Filter '*.csv' |
        Where-Object {

            # Do NOT accidentally use our own output files.
            $_.Name -notlike 'Lost_Device_Audit_*' -and
            $_.Name -notlike 'Lost_Device_Full_Report_*' -and
            $_.Name -notlike '~$*'

        } |
        Sort-Object LastWriteTime -Descending
    )

    if ($files.Count -eq 0) {
        throw "No usable input CSV was found in: $Folder"
    }

    return $files[0].FullName
}


function Find-CsvColumn {
    param(
        [string[]]$Headers,
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {

        foreach ($header in $Headers) {

            if ($header -ieq $candidate) {
                return $header
            }
        }
    }

    return $null
}


function Read-InputCsv {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Input CSV not found: $Path"
    }

    if ([IO.Path]::GetExtension($Path) -ine '.csv') {
        throw "Input file must be a CSV: $Path"
    }

    Write-Host ""
    Write-Host "Reading CSV..." -ForegroundColor Cyan

    $csvRows = @(Import-Csv -LiteralPath $Path)

    if ($csvRows.Count -eq 0) {
        throw "The input CSV contains no data rows."
    }


    # --------------------------------------------------------
    # GET CSV HEADERS
    # --------------------------------------------------------

    $headers = @(
        $csvRows[0].PSObject.Properties.Name
    )

    Write-Host ""
    Write-Host "CSV columns found:" -ForegroundColor DarkCyan

    foreach ($header in $headers) {
        Write-Host "  $header" -ForegroundColor Gray
    }


    # --------------------------------------------------------
    # FIND DEVICE COLUMN
    # --------------------------------------------------------

    $deviceColumn = Find-CsvColumn `
        -Headers $headers `
        -Candidates @(

            'Device',
            'Device Name',
            'DeviceName',

            'Computer',
            'Computer Name',
            'ComputerName',

            'Hostname',
            'Host Name',
            'HostName',

            'Workstation',
            'Workstation Name',
            'WorkstationName',

            'Machine',
            'Machine Name',
            'MachineName',

            'PC',
            'PC Name',
            'PCName',

            'Asset Name',
            'Endpoint',

            'Name'
        )


    # --------------------------------------------------------
    # FIND OPTIONAL OWNER COLUMN
    # --------------------------------------------------------

    $ownerColumn = Find-CsvColumn `
        -Headers $headers `
        -Candidates @(

            'User / Owner',
            'User/Owner',

            'UserOwner',

            'Owner',
            'Owner Name',
            'OwnerName',

            'User',
            'Username',
            'User Name',

            'Assigned User',
            'AssignedUser',

            'Last User',
            'LastUser',

            'Primary User',
            'PrimaryUser',

            'Logged On User',
            'LoggedOnUser'
        )


    # --------------------------------------------------------
    # FIND OPTIONAL MANAGER COLUMN
    # --------------------------------------------------------

    $managerColumn = Find-CsvColumn `
        -Headers $headers `
        -Candidates @(

            'Manager',
            'Manager Name',
            'ManagerName',

            'Supervisor',
            'Supervisor Name',

            'Owner Manager'
        )


    # --------------------------------------------------------
    # FIND OPTIONAL NOTES COLUMN
    # --------------------------------------------------------

    $notesColumn = Find-CsvColumn `
        -Headers $headers `
        -Candidates @(

            'Source Notes',
            'SourceNotes',

            'Notes',
            'Note',

            'Comments',
            'Comment'
        )


    if ([string]::IsNullOrWhiteSpace($deviceColumn)) {

        $available = $headers -join ', '

        throw @"
Could not identify the DEVICE column.

Available CSV headers:
$available

Expected something like:
Device
Device Name
Computer Name
Hostname
Workstation
Workstation Name
Machine Name
Name
"@
    }


    # --------------------------------------------------------
    # SHOW EXACT COLUMNS BEING USED
    # --------------------------------------------------------

    Write-Host ""
    Write-Host "COLUMN MAPPING" -ForegroundColor Yellow
    Write-Host "----------------------------------------"

    Write-Host "Device column : $deviceColumn" -ForegroundColor Green

    if ($ownerColumn) {
        Write-Host "Owner column  : $ownerColumn" -ForegroundColor Green
    }
    else {
        Write-Host "Owner column  : NOT FOUND - SCCM will be used" `
            -ForegroundColor DarkYellow
    }

    if ($managerColumn) {
        Write-Host "Manager column: $managerColumn" -ForegroundColor Green
    }
    else {
        Write-Host "Manager column: NOT FOUND - AD manager will be used" `
            -ForegroundColor DarkYellow
    }

    if ($notesColumn) {
        Write-Host "Notes column  : $notesColumn" -ForegroundColor Green
    }
    else {
        Write-Host "Notes column  : NOT FOUND" `
            -ForegroundColor DarkGray
    }


    # --------------------------------------------------------
    # NORMALIZE DATA
    # --------------------------------------------------------

    $results = New-Object System.Collections.Generic.List[object]

    $seen = @{}

    $duplicateCount = 0
    $blankCount = 0

    foreach ($row in $csvRows) {

        $device = [string](
            Get-SafePropertyValue `
                -Object $row `
                -Name $deviceColumn
        )

        $device = $device.Trim()

        if ([string]::IsNullOrWhiteSpace($device)) {

            $blankCount++

            continue
        }


        $key = $device.ToUpperInvariant()


        # Ignore duplicates.
        if ($seen.ContainsKey($key)) {

            $duplicateCount++

            continue
        }

        $seen[$key] = $true


        # Owner
        $owner = ''

        if ($ownerColumn) {

            $owner = [string](
                Get-SafePropertyValue `
                    -Object $row `
                    -Name $ownerColumn
            )

            $owner = $owner.Trim()
        }


        # Manager
        $manager = ''

        if ($managerColumn) {

            $manager = [string](
                Get-SafePropertyValue `
                    -Object $row `
                    -Name $managerColumn
            )

            $manager = $manager.Trim()
        }


        # Notes
        $notes = ''

        if ($notesColumn) {

            $notes = [string](
                Get-SafePropertyValue `
                    -Object $row `
                    -Name $notesColumn
            )

            $notes = $notes.Trim()
        }


        $results.Add(
            [pscustomobject]@{

                Manager     = $manager

                UserOwner   = $owner

                Device      = $device

                SourceNotes = $notes
            }
        )
    }


    if ($results.Count -eq 0) {
        throw "No valid device names were found in the CSV."
    }


    Write-Host ""
    Write-Host "CSV SUMMARY" -ForegroundColor Yellow
    Write-Host "----------------------------------------"

    Write-Host "Rows in CSV:       $($csvRows.Count)"
    Write-Host "Unique devices:    $($results.Count)" `
        -ForegroundColor Green

    Write-Host "Duplicate devices: $duplicateCount"

    Write-Host "Blank device rows: $blankCount"

    Write-Host ""

    return $results.ToArray()
}


# ============================================================
# SCCM
# ============================================================

function Initialize-SCCM {
    param(
        [string]$RequestedSiteCode,
        [string]$RequestedSiteServer
    )


    # --------------------------------------------------------
    # TRY INSTALLED MODULE FIRST
    # --------------------------------------------------------

    $module = Get-Module `
        ConfigurationManager `
        -ListAvailable |
        Select-Object -First 1


    # --------------------------------------------------------
    # TRY SCCM CONSOLE PATH
    # --------------------------------------------------------

    if (-not $module -and $env:SMS_ADMIN_UI_PATH) {

        $candidate = Join-Path `
            (Split-Path $env:SMS_ADMIN_UI_PATH -Parent) `
            'ConfigurationManager.psd1'

        if (Test-Path $candidate) {

            Import-Module `
                $candidate `
                -ErrorAction Stop
        }
    }
    elseif ($module) {

        Import-Module `
            $module.Path `
            -ErrorAction Stop
    }


    if (-not (Get-Module ConfigurationManager)) {

        throw @"
ConfigurationManager PowerShell module was not found.

Run this from a computer with the SCCM / Configuration Manager
console installed.
"@
    }


    # --------------------------------------------------------
    # FIND EXISTING SCCM DRIVE
    # --------------------------------------------------------

    $cmDrive = Get-PSDrive `
        -PSProvider CMSite `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1


    # --------------------------------------------------------
    # CREATE DRIVE IF SITE INFO WAS PROVIDED
    # --------------------------------------------------------

    if (
        -not $cmDrive -and
        $RequestedSiteCode -and
        $RequestedSiteServer
    ) {

        New-PSDrive `
            -Name $RequestedSiteCode `
            -PSProvider CMSite `
            -Root $RequestedSiteServer `
            -ErrorAction Stop |
        Out-Null


        $cmDrive = Get-PSDrive `
            -Name $RequestedSiteCode `
            -PSProvider CMSite `
            -ErrorAction Stop
    }


    if (-not $cmDrive) {

        throw @"
No SCCM CMSite PowerShell drive was found.

Either:

1. Open the Configuration Manager PowerShell console first

OR

2. Run the script with:

-SiteCode "ABC" -SiteServer "SCCM01.contoso.local"
"@
    }


    return $cmDrive
}


# ============================================================
# DEVICE RECOMMENDATION
# ============================================================

function Get-Recommendation {
    param(
        [bool]$PingOnline,

        [bool]$ADFound,

        $ADEnabled,

        [bool]$SCCMFound,

        $SCCMActive,

        $SCCMIsClient
    )


    if ($PingOnline) {

        return 'Locate - Online now'
    }


    if (
        $ADFound -and
        $ADEnabled -eq $true
    ) {

        return 'Locate - Active in AD'
    }


    if (
        $SCCMFound -and
        (
            $SCCMActive -eq $true -or
            $SCCMIsClient -eq $true
        )
    ) {

        return 'Locate - Active in SCCM'
    }


    if (
        -not $ADFound -and
        -not $SCCMFound
    ) {

        return 'Retirement Review - Not in AD or SCCM'
    }


    if (
        $ADFound -and
        $ADEnabled -eq $false -and
        (
            -not $SCCMFound -or
            $SCCMActive -eq $false
        )
    ) {

        return 'Retirement Review - AD disabled / SCCM inactive'
    }


    return 'Needs Review'
}


# ============================================================
# AD USER / OWNER RESOLUTION
# ============================================================

function Normalize-UserCandidate {
    param(
        [string]$Value
    )


    if ([string]::IsNullOrWhiteSpace($Value)) {

        return ''
    }


    $valueClean = $Value.Trim()


    # DOMAIN\username
    if ($valueClean -match '\\') {

        $valueClean = ($valueClean -split '\\')[-1]
    }


    # username@domain.com
    if ($valueClean -match '@') {

        $valueClean = ($valueClean -split '@')[0]
    }


    return $valueClean.Trim()
}


function Resolve-ADUserCandidate {
    param(
        [string]$Candidate
    )


    $original = $Candidate

    $candidateClean = Normalize-UserCandidate `
        -Value $Candidate


    if ([string]::IsNullOrWhiteSpace($candidateClean)) {

        return $null
    }


    $escaped = Escape-LdapValue `
        -Value $candidateClean


    $displayEscaped = Escape-LdapValue `
        -Value $original.Trim()


    # --------------------------------------------------------
    # TRY DIRECT IDENTITY
    # --------------------------------------------------------

    try {

        $user = Get-ADUser `
            -Identity $candidateClean `
            -Properties `
                GivenName,
                Surname,
                DisplayName,
                Manager,
                EmployeeID,
                SamAccountName `
            -ErrorAction Stop


        if ($user) {

            return $user
        }
    }
    catch {}


    # --------------------------------------------------------
    # TRY SAM / EMPLOYEE ID / UPN / DISPLAY NAME
    # --------------------------------------------------------

    try {

        $ldap = @"
(|(sAMAccountName=$escaped)(employeeID=$escaped)(userPrincipalName=$escaped@*)(displayName=$displayEscaped))
"@


        $user = Get-ADUser `
            -LDAPFilter $ldap `
            -Properties `
                GivenName,
                Surname,
                DisplayName,
                Manager,
                EmployeeID,
                SamAccountName `
            -ErrorAction Stop |
        Select-Object -First 1


        if ($user) {

            return $user
        }
    }
    catch {}


    # --------------------------------------------------------
    # TRY EXACT DISPLAY NAME
    # --------------------------------------------------------

    if ($original -match '\s') {

        try {

            $safeName = $original.Replace(
                "'",
                "''"
            )


            $user = Get-ADUser `
                -Filter "DisplayName -eq '$safeName'" `
                -Properties `
                    GivenName,
                    Surname,
                    DisplayName,
                    Manager,
                    EmployeeID,
                    SamAccountName `
                -ErrorAction Stop |
            Select-Object -First 1


            if ($user) {

                return $user
            }
        }
        catch {}
    }


    return $null
}


function Get-ManagerDisplayName {
    param(
        $User
    )


    if (
        $null -eq $User -or
        [string]::IsNullOrWhiteSpace(
            [string]$User.Manager
        )
    ) {

        return ''
    }


    try {

        $manager = Get-ADUser `
            -Identity $User.Manager `
            -Properties DisplayName `
            -ErrorAction Stop


        return [string]$manager.DisplayName
    }
    catch {

        return ''
    }
}


function Resolve-OwnerInfo {
    param(
        [string]$SourceOwner,

        [string]$SCCMPrimaryUser,

        [hashtable]$Cache
    )


    $candidates = New-Object `
        System.Collections.Generic.List[string]


    # Prefer SCCM user.
    if (
        -not [string]::IsNullOrWhiteSpace(
            $SCCMPrimaryUser
        )
    ) {

        $candidates.Add($SCCMPrimaryUser)
    }


    # Fall back to CSV owner.
    if (
        -not [string]::IsNullOrWhiteSpace(
            $SourceOwner
        )
    ) {

        $candidates.Add($SourceOwner)
    }


    foreach ($candidate in $candidates) {

        $cacheKey = $candidate.Trim().ToUpperInvariant()


        # ----------------------------------------------------
        # CACHE LOOKUP
        # ----------------------------------------------------

        if ($Cache.ContainsKey($cacheKey)) {

            $cached = $Cache[$cacheKey]

            if ($null -ne $cached) {

                return $cached
            }

            continue
        }


        # ----------------------------------------------------
        # AD LOOKUP
        # ----------------------------------------------------

        $user = Resolve-ADUserCandidate `
            -Candidate $candidate


        if ($user) {

            $resolved = [pscustomobject]@{

                FirstName = [string]$user.GivenName

                LastName = [string]$user.Surname

                DisplayName = [string]$user.DisplayName

                Account = [string]$user.SamAccountName

                Manager = Get-ManagerDisplayName `
                    -User $user

                LookupStatus = 'Resolved in AD'
            }


            $Cache[$cacheKey] = $resolved

            return $resolved
        }


        $Cache[$cacheKey] = $null
    }


    # --------------------------------------------------------
    # PRESERVE OBVIOUS FULL NAMES
    # BUT DO NOT CLAIM THEY WERE RESOLVED IN AD
    # --------------------------------------------------------

    if (
        $SourceOwner -match '^\s*([^\d\s]+)\s+(.+?)\s*$' -and
        $SourceOwner -notmatch '^[A-Za-z]{0,4}\d'
    ) {

        $parts = $SourceOwner.Trim() -split '\s+'


        if ($parts.Count -ge 2) {

            return [pscustomobject]@{

                FirstName = $parts[0]

                LastName = (
                    $parts[1..($parts.Count - 1)] -join ' '
                )

                DisplayName = $SourceOwner.Trim()

                Account = ''

                Manager = ''

                LookupStatus = 'Source name - AD not resolved'
            }
        }
    }


    # --------------------------------------------------------
    # UNRESOLVED
    # --------------------------------------------------------

    $fallbackAccount = ''

    if ($SCCMPrimaryUser) {

        $fallbackAccount = Normalize-UserCandidate `
            -Value $SCCMPrimaryUser
    }
    elseif ($SourceOwner) {

        $fallbackAccount = Normalize-UserCandidate `
            -Value $SourceOwner
    }


    return [pscustomobject]@{

        FirstName = ''

        LastName = ''

        DisplayName = ''

        Account = $fallbackAccount

        Manager = ''

        LookupStatus = 'Owner not resolved'
    }
}


# ============================================================
# EXCEL OUTPUT HELPERS
# ============================================================

function Set-HeaderStyle {
    param(
        $Range
    )


    $Range.Font.Bold = $true

    $Range.Interior.Color = 7874585

    $Range.Font.Color = 16777215
}


function Write-WorksheetData {
    param(
        $Sheet,

        [object[]]$Data,

        [string[]]$Headers
    )


    # --------------------------------------------------------
    # HEADERS
    # --------------------------------------------------------

    for (
        $column = 0;
        $column -lt $Headers.Count;
        $column++
    ) {

        $Sheet.Cells.Item(
            1,
            $column + 1
        ).Value2 = [string]$Headers[$column]
    }


    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    $rowNumber = 2


    foreach ($item in $Data) {

        for (
            $column = 0;
            $column -lt $Headers.Count;
            $column++
        ) {

            $value = Get-SafePropertyValue `
                -Object $item `
                -Name $Headers[$column]


            if ($value -is [datetime]) {

                $text = $value.ToString(
                    'M/d/yyyy h:mm tt'
                )
            }
            else {

                $text = Convert-ToDisplayText `
                    -Value $value
            }


            $Sheet.Cells.Item(
                $rowNumber,
                $column + 1
            ).Value2 = [string]$text
        }


        $rowNumber++
    }


    # --------------------------------------------------------
    # FORMAT
    # --------------------------------------------------------

    $lastRow = [Math]::Max(
        2,
        $rowNumber - 1
    )

    $lastColumn = $Headers.Count


    $range = $Sheet.Range(
        $Sheet.Cells.Item(1, 1),
        $Sheet.Cells.Item(
            $lastRow,
            $lastColumn
        )
    )


    Set-HeaderStyle `
        -Range $Sheet.Range(
            $Sheet.Cells.Item(1, 1),
            $Sheet.Cells.Item(
                1,
                $lastColumn
            )
        )


    $range.AutoFilter() | Out-Null

    $range.VerticalAlignment = -4160

    $range.Columns.AutoFit() | Out-Null


    for (
        $column = 1;
        $column -le $lastColumn;
        $column++
    ) {

        if (
            $Sheet.Columns.Item(
                $column
            ).ColumnWidth -gt 32
        ) {

            $Sheet.Columns.Item(
                $column
            ).ColumnWidth = 32
        }
    }


    # Freeze header row.
    $Sheet.Activate()

    $Sheet.Application.ActiveWindow.SplitRow = 1

    $Sheet.Application.ActiveWindow.FreezePanes = $true
}


# ============================================================
# CREATE FINAL XLSX
# ============================================================

function Write-TrimmedWorkbook {
    param(
        [string]$Path,

        [object[]]$Results,

        [string]$SourcePath
    )


    $excel = $null
    $book = $null

    $dashboard = $null
    $byManager = $null
    $online = $null
    $offline = $null
    $retire = $null
    $managerSummary = $null


    try {

        Write-Host ""
        Write-Host "Starting Excel..." -ForegroundColor Cyan


        $excel = New-Object `
            -ComObject Excel.Application


        $excel.Visible = $false

        $excel.DisplayAlerts = $false


        $book = $excel.Workbooks.Add()


        while ($book.Worksheets.Count -lt 6) {

            [void]$book.Worksheets.Add()
        }


        # ----------------------------------------------------
        # SHEETS
        # ----------------------------------------------------

        $dashboard = $book.Worksheets.Item(1)

        $dashboard.Name = 'Dashboard'


        $byManager = $book.Worksheets.Item(2)

        $byManager.Name = 'By Manager'


        $online = $book.Worksheets.Item(3)

        $online.Name = 'Online'


        $offline = $book.Worksheets.Item(4)

        $offline.Name = 'Offline'


        $retire = $book.Worksheets.Item(5)

        $retire.Name = 'Retirement Review'


        $managerSummary = $book.Worksheets.Item(6)

        $managerSummary.Name = 'Manager Summary'


        # ----------------------------------------------------
        # WORKING SHEET COLUMNS
        # ----------------------------------------------------

        $headers = @(

            'Manager',

            'Owner First Name',

            'Owner Last Name',

            'Owner Lookup',

            'Device',

            'Online Status',

            'Ping IP',

            'AD Enabled',

            'AD Last Logon',

            'SCCM Active',

            'SCCM Last Active',

            'SCCM Primary User',

            'Recommendation',

            'Team Notes'
        )


        # ----------------------------------------------------
        # CONVERT TECHNICAL RESULTS TO TEAM-FRIENDLY ROWS
        # ----------------------------------------------------

        $teamRows = @(

            $Results |

            ForEach-Object {

                [pscustomobject][ordered]@{

                    'Manager' =
                        $_.Manager

                    'Owner First Name' =
                        $_.OwnerFirstName

                    'Owner Last Name' =
                        $_.OwnerLastName

                    'Owner Lookup' =
                        $_.OwnerLookup

                    'Device' =
                        $_.Device

                    'Online Status' =
                        if ($_.PingOnline) {
                            'Online'
                        }
                        else {
                            'Offline'
                        }

                    'Ping IP' =
                        $_.PingIP

                    'AD Enabled' =
                        $_.ADEnabled

                    'AD Last Logon' =
                        $_.ADLastLogon

                    'SCCM Active' =
                        $_.SCCMActive

                    'SCCM Last Active' =
                        $_.SCCMLastActive

                    'SCCM Primary User' =
                        $_.SCCMPrimaryUser

                    'Recommendation' =
                        $_.Recommendation

                    'Team Notes' = ''
                }
            } |

            Sort-Object `
                Manager,
                'Owner Last Name',
                'Owner First Name',
                Device
        )


        # ----------------------------------------------------
        # WRITE WORKSHEETS
        # ----------------------------------------------------

        Write-WorksheetData `
            -Sheet $byManager `
            -Data $teamRows `
            -Headers $headers


        Write-WorksheetData `
            -Sheet $online `
            -Data @(
                $teamRows |
                Where-Object {
                    $_.'Online Status' -eq 'Online'
                }
            ) `
            -Headers $headers


        Write-WorksheetData `
            -Sheet $offline `
            -Data @(
                $teamRows |
                Where-Object {
                    $_.'Online Status' -eq 'Offline'
                }
            ) `
            -Headers $headers


        Write-WorksheetData `
            -Sheet $retire `
            -Data @(
                $teamRows |
                Where-Object {
                    $_.Recommendation -like 'Retirement Review*'
                }
            ) `
            -Headers $headers


        # ----------------------------------------------------
        # ONLINE / OFFLINE CELL COLORS
        # ----------------------------------------------------

        foreach (
            $worksheet in @(
                $byManager,
                $online,
                $offline,
                $retire
            )
        ) {

            $last = $worksheet.UsedRange.Rows.Count


            for (
                $row = 2;
                $row -le $last;
                $row++
            ) {

                $status = [string](
                    $worksheet.Cells.Item(
                        $row,
                        6
                    ).Value2
                )


                if ($status -eq 'Online') {

                    $worksheet.Cells.Item(
                        $row,
                        6
                    ).Interior.Color = 13434828
                }


                if ($status -eq 'Offline') {

                    $worksheet.Cells.Item(
                        $row,
                        6
                    ).Interior.Color = 13551615
                }
            }
        }


        # ====================================================
        # DASHBOARD
        # ====================================================

        $dashboard.Range(
            'A1:F1'
        ).Merge()


        $dashboard.Cells.Item(
            1,
            1
        ).Value2 = 'Lost Device Audit Dashboard'


        $dashboard.Cells.Item(
            1,
            1
        ).Font.Bold = $true


        $dashboard.Cells.Item(
            1,
            1
        ).Font.Size = 16


        $onlineCount = @(
            $Results |
            Where-Object {
                $_.PingOnline
            }
        ).Count


        $offlineCount =
            $Results.Count -
            $onlineCount


        $retireCount = @(
            $Results |
            Where-Object {
                $_.Recommendation -like 'Retirement Review*'
            }
        ).Count


        $ownerMissing = @(
            $Results |
            Where-Object {
                $_.OwnerLookup -eq 'Owner not resolved'
            }
        ).Count


        $managerCount = @(
            $Results.Manager |
            Where-Object {
                $_ -and
                $_ -ne 'Manager Not Identified'
            } |
            Sort-Object -Unique
        ).Count


        $dashData = @(

            @(
                'Metric',
                'Value'
            ),

            @(
                'Total Devices',
                $Results.Count
            ),

            @(
                'Online',
                $onlineCount
            ),

            @(
                'Offline',
                $offlineCount
            ),

            @(
                'Retirement Review',
                $retireCount
            ),

            @(
                'Owner Not Resolved',
                $ownerMissing
            ),

            @(
                'Managers',
                $managerCount
            ),

            @(
                'Source CSV',
                $SourcePath
            ),

            @(
                'Audit Date',
                (Get-Date).ToString(
                    'M/d/yyyy h:mm tt'
                )
            )
        )


        for (
            $row = 0;
            $row -lt $dashData.Count;
            $row++
        ) {

            $dashboard.Cells.Item(
                $row + 3,
                1
            ).Value2 = [string]$dashData[$row][0]


            $dashboard.Cells.Item(
                $row + 3,
                2
            ).Value2 = [string]$dashData[$row][1]
        }


        Set-HeaderStyle `
            -Range $dashboard.Range(
                'A3:B3'
            )


        $dashboard.Columns.Item(
            1
        ).ColumnWidth = 24


        $dashboard.Columns.Item(
            2
        ).ColumnWidth = 60


        # ====================================================
        # MANAGER SUMMARY
        # ====================================================

        $managerSummary.Cells.Item(
            1,
            1
        ).Value2 = 'Manager'


        $managerSummary.Cells.Item(
            1,
            2
        ).Value2 = 'Total'


        $managerSummary.Cells.Item(
            1,
            3
        ).Value2 = 'Online'


        $managerSummary.Cells.Item(
            1,
            4
        ).Value2 = 'Offline'


        Set-HeaderStyle `
            -Range $managerSummary.Range(
                'A1:D1'
            )


        $managerNames = @(

            $Results.Manager |

            ForEach-Object {

                if (
                    [string]::IsNullOrWhiteSpace(
                        $_
                    )
                ) {

                    'Manager Not Identified'
                }
                else {

                    $_
                }
            } |

            Sort-Object -Unique
        )


        $summaryRow = 2


        foreach ($manager in $managerNames) {

            $managerResults = @(

                $Results |

                Where-Object {

                    $currentManager = $_.Manager


                    if (
                        [string]::IsNullOrWhiteSpace(
                            $currentManager
                        )
                    ) {

                        $currentManager =
                            'Manager Not Identified'
                    }


                    $currentManager -eq $manager
                }
            )


            $managerOnline = @(

                $managerResults |

                Where-Object {
                    $_.PingOnline
                }

            ).Count


            $managerSummary.Cells.Item(
                $summaryRow,
                1
            ).Value2 = [string]$manager


            $managerSummary.Cells.Item(
                $summaryRow,
                2
            ).Value2 = [string]$managerResults.Count


            $managerSummary.Cells.Item(
                $summaryRow,
                3
            ).Value2 = [string]$managerOnline


            $managerSummary.Cells.Item(
                $summaryRow,
                4
            ).Value2 = [string](
                $managerResults.Count -
                $managerOnline
            )


            $summaryRow++
        }


        $managerSummary.Range(
            "A1:D$($summaryRow - 1)"
        ).AutoFilter() |
        Out-Null


        $managerSummary.Columns.AutoFit() |
        Out-Null


        # ====================================================
        # SAVE
        # ====================================================

        $folder = Split-Path `
            $Path `
            -Parent


        if (
            $folder -and
            -not (Test-Path $folder)
        ) {

            New-Item `
                -ItemType Directory `
                -Path $folder `
                -Force |
            Out-Null
        }


        Write-Host ""
        Write-Host "Saving Excel workbook..." `
            -ForegroundColor Cyan


        # XLSX
        $book.SaveAs(
            $Path,
            51
        )
    }
    finally {

        if ($book) {

            $book.Close(
                $false
            ) |
            Out-Null
        }


        if ($excel) {

            $excel.Quit()
        }


        foreach (
            $object in @(
                $dashboard,
                $byManager,
                $online,
                $offline,
                $retire,
                $managerSummary,
                $book,
                $excel
            )
        ) {

            if ($object) {

                try {

                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $object
                    )
                }
                catch {}
            }
        }


        [GC]::Collect()

        [GC]::WaitForPendingFinalizers()
    }
}


# ============================================================
# MAIN
# ============================================================

Write-Host ""
Write-Host "==============================================" `
    -ForegroundColor Cyan

Write-Host " LOST DEVICE AUDIT" `
    -ForegroundColor Cyan

Write-Host " CSV -> AD + SCCM + ONE PING -> XLSX" `
    -ForegroundColor Cyan

Write-Host "==============================================" `
    -ForegroundColor Cyan

Write-Host ""


# ============================================================
# FIND INPUT CSV
# ============================================================

if (
    [string]::IsNullOrWhiteSpace(
        $InputCsv
    )
) {

    $InputCsv = Get-LatestInputCsv `
        -Folder $PSScriptRoot
}


$InputCsv = (
    Resolve-Path `
        -LiteralPath $InputCsv
).Path


Write-Host "INPUT FILE" -ForegroundColor Yellow

Write-Host "----------------------------------------"

Write-Host $InputCsv `
    -ForegroundColor Green

Write-Host ""


# ============================================================
# OUTPUT FILE
# ============================================================

if (
    [string]::IsNullOrWhiteSpace(
        $OutputWorkbook
    )
) {

    $OutputWorkbook = Join-Path `
        $PSScriptRoot `
        (
            "Lost_Device_Audit_{0}.xlsx" -f
            (
                Get-Date `
                    -Format 'yyyyMMdd_HHmm'
            )
        )
}


# ============================================================
# READ CSV
# ============================================================

$devices = @(
    Read-InputCsv `
        -Path $InputCsv
)


if ($devices.Count -eq 0) {

    throw 'No devices were found in the input CSV.'
}


# ============================================================
# ACTIVE DIRECTORY
# ============================================================

Write-Host ""
Write-Host "Loading Active Directory module..." `
    -ForegroundColor Cyan


Import-Module `
    ActiveDirectory `
    -ErrorAction Stop


Write-Host "Active Directory module loaded." `
    -ForegroundColor Green


# ============================================================
# SCCM
# ============================================================

Write-Host ""
Write-Host "Connecting to Configuration Manager..." `
    -ForegroundColor Cyan


$originalLocation = Get-Location


$cmDrive = Initialize-SCCM `
    -RequestedSiteCode $SiteCode `
    -RequestedSiteServer $SiteServer


$resolvedSiteCode = $cmDrive.Name


Write-Host (
    "Using SCCM site {0} ({1})" -f
    $resolvedSiteCode,
    $cmDrive.Root
) -ForegroundColor Green


# ============================================================
# LOAD SCCM DEVICE RECORDS ONCE
# ============================================================

Write-Host ""
Write-Host "Loading matching SCCM records..." `
    -ForegroundColor Cyan


$wanted = @{}


foreach ($deviceEntry in $devices) {

    $wanted[
        $deviceEntry.Device.ToUpperInvariant()
    ] = $true
}


$sccmMap = @{}


try {

    Set-Location "$resolvedSiteCode`:"


    Get-CMDevice -Fast |

    ForEach-Object {

        $name = [string](
            Get-SafePropertyValue `
                -Object $_ `
                -Name 'Name'
        )


        if (
            $name -and
            $wanted.ContainsKey(
                $name.ToUpperInvariant()
            )
        ) {

            $sccmMap[
                $name.ToUpperInvariant()
            ] = $_
        }
    }
}
finally {

    Set-Location $originalLocation
}


Write-Host (
    "Matched {0} of {1} devices in SCCM." -f
    $sccmMap.Count,
    $devices.Count
) -ForegroundColor Green


# ============================================================
# LOAD AD COMPUTER RECORDS
# ============================================================

Write-Host ""
Write-Host "Loading matching AD computer records..." `
    -ForegroundColor Cyan


$adMap = @{}


$names = @(

    $devices |

    ForEach-Object {
        $_.Device
    } |

    Sort-Object -Unique
)


# Keep LDAP queries reasonably sized.
$chunkSize = 75


for (
    $index = 0;
    $index -lt $names.Count;
    $index += $chunkSize
) {

    $end = [Math]::Min(
        $index + $chunkSize - 1,
        $names.Count - 1
    )


    $chunk = @(
        $names[$index..$end]
    )


    $parts = @(

        $chunk |

        ForEach-Object {

            $escapedName = Escape-LdapValue `
                -Value $_


            "(name=$escapedName)"
        }
    )


    $ldap = '(|' +
        ($parts -join '') +
        ')'


    try {

        Get-ADComputer -LDAPFilter $ldap -Properties Enabled,LastLogonDate |
    ForEach-Object {
        $adMap[$_.Name.ToUpperInvariant()] = $_
    }
    }
    catch {

        # Fallback to individual lookups if the batch query fails.
        foreach ($name in $chunk) {

            try {

                $computer = Get-ADComputer `
                    -Identity $name `
                    -Properties `
                        Enabled,
                        LastLogonDate `
                    -ErrorAction Stop


                $adMap[
                    $name.ToUpperInvariant()
                ] = $computer
            }
            catch {}
        }
    }
}


Write-Host (
    "Matched {0} of {1} devices in AD." -f
    $adMap.Count,
    $devices.Count
) -ForegroundColor Green


# ============================================================
# DEVICE CHECK
# ============================================================

Write-Host ""
Write-Host "Checking devices..." `
    -ForegroundColor Cyan

Write-Host "ONE sequential ping per device." `
    -ForegroundColor Yellow

Write-Host "No retries. No parallel pinging." `
    -ForegroundColor Yellow

Write-Host ""


$results = New-Object `
    System.Collections.Generic.List[object]


$ownerCache = @{}


$pingObject = New-Object `
    System.Net.NetworkInformation.Ping


$currentIndex = 0

$total = $devices.Count


try {

    foreach ($entry in $devices) {

        $currentIndex++


        $device = $entry.Device.Trim()


        $key = $device.ToUpperInvariant()


        Write-Progress `
            -Activity 'Lost device audit - one ping at a time' `
            -Status "$currentIndex of $total - $device" `
            -PercentComplete (
                ($currentIndex / $total) *
                100
            )


        # ====================================================
        # ONE PING ONLY
        # ====================================================

        $pingOnline = $false

        $pingIP = ''


        try {

            $reply = $pingObject.Send(
                $device,
                $PingTimeoutMs
            )


            if (
                $reply -and
                $reply.Status -eq
                    [System.Net.NetworkInformation.IPStatus]::Success
            ) {

                $pingOnline = $true

                $pingIP = [string]$reply.Address
            }
        }
        catch {

            # No retry.
            $pingOnline = $false

            $pingIP = ''
        }


        # ====================================================
        # AD COMPUTER
        # ====================================================

        if ($adMap.ContainsKey($key)) {

            $ad = $adMap[$key]
        }
        else {

            $ad = $null
        }


        $adFound = (
            $null -ne $ad
        )


        if ($ad) {

            $adEnabled = $ad.Enabled

            $adLastLogon = $ad.LastLogonDate
        }
        else {

            $adEnabled = $null

            $adLastLogon = $null
        }


        # ====================================================
        # SCCM DEVICE
        # ====================================================

        if ($sccmMap.ContainsKey($key)) {

            $cm = $sccmMap[$key]
        }
        else {

            $cm = $null
        }


        $sccmFound = (
            $null -ne $cm
        )


        if ($cm) {

            $sccmIsClient =
                Get-SafePropertyValue `
                    -Object $cm `
                    -Name 'IsClient'


            $sccmActive =
                Get-SafePropertyValue `
                    -Object $cm `
                    -Name 'IsActive'


            $sccmLastActive =
                Get-SafePropertyValue `
                    -Object $cm `
                    -Name 'LastActiveTime'


            $sccmPrimaryUser =
                Convert-ToDisplayText `
                    -Value (
                        Get-SafePropertyValue `
                            -Object $cm `
                            -Name 'UserName'
                    )
        }
        else {

            $sccmIsClient = $null

            $sccmActive = $null

            $sccmLastActive = $null

            $sccmPrimaryUser = ''
        }


        # ====================================================
        # OWNER
        # ====================================================

        $owner = Resolve-OwnerInfo `
            -SourceOwner $entry.UserOwner `
            -SCCMPrimaryUser $sccmPrimaryUser `
            -Cache $ownerCache


        # ====================================================
        # MANAGER
        # ====================================================

        $manager = $owner.Manager


        # Fall back to CSV manager.
        if (
            [string]::IsNullOrWhiteSpace(
                $manager
            )
        ) {

            $manager = $entry.Manager
        }


        if (
            [string]::IsNullOrWhiteSpace(
                $manager
            )
        ) {

            $manager =
                'Manager Not Identified'
        }


        # ====================================================
        # RECOMMENDATION
        # ====================================================

        $recommendation = Get-Recommendation `
            -PingOnline $pingOnline `
            -ADFound $adFound `
            -ADEnabled $adEnabled `
            -SCCMFound $sccmFound `
            -SCCMActive $sccmActive `
            -SCCMIsClient $sccmIsClient


        # ====================================================
        # RESULT
        # ====================================================

        $results.Add(

            [pscustomobject]@{

                Manager =
                    $manager

                OwnerFirstName =
                    $owner.FirstName

                OwnerLastName =
                    $owner.LastName

                OwnerDisplayName =
                    $owner.DisplayName

                OwnerAccount =
                    $owner.Account

                OwnerLookup =
                    $owner.LookupStatus

                SourceOwner =
                    $entry.UserOwner

                Device =
                    $device

                PingOnline =
                    $pingOnline

                PingIP =
                    $pingIP

                ADFound =
                    $adFound

                ADEnabled =
                    $adEnabled

                ADLastLogon =
                    $adLastLogon

                SCCMFound =
                    $sccmFound

                SCCMIsClient =
                    $sccmIsClient

                SCCMActive =
                    $sccmActive

                SCCMLastActive =
                    $sccmLastActive

                SCCMPrimaryUser =
                    $sccmPrimaryUser

                Recommendation =
                    $recommendation

                SourceNotes =
                    $entry.SourceNotes

                CheckedAt =
                    Get-Date
            }
        )
    }
}
finally {

    if ($pingObject) {

        $pingObject.Dispose()
    }


    Write-Progress `
        -Activity 'Lost device audit - one ping at a time' `
        -Completed
}


# ============================================================
# SAVE TECHNICAL CSV
# ============================================================

$checkpoint = [IO.Path]::ChangeExtension(
    $OutputWorkbook,
    '.csv'
)


Write-Host ""
Write-Host "Saving technical CSV:" `
    -ForegroundColor Cyan

Write-Host $checkpoint `
    -ForegroundColor Gray


$results.ToArray() |

Export-Csv `
    -Path $checkpoint `
    -NoTypeInformation `
    -Encoding UTF8


# ============================================================
# CREATE XLSX
# ============================================================

Write-Host ""
Write-Host "Creating Excel workbook:" `
    -ForegroundColor Cyan

Write-Host $OutputWorkbook `
    -ForegroundColor Gray


Write-TrimmedWorkbook `
    -Path $OutputWorkbook `
    -Results $results.ToArray() `
    -SourcePath $InputCsv


# ============================================================
# FINAL SUMMARY
# ============================================================

$onlineTotal = @(

    $results |

    Where-Object {
        $_.PingOnline
    }

).Count


$offlineTotal =
    $results.Count -
    $onlineTotal


$ownersUnresolved = @(

    $results |

    Where-Object {
        $_.OwnerLookup -eq
            'Owner not resolved'
    }

).Count


$retirementTotal = @(

    $results |

    Where-Object {
        $_.Recommendation -like
            'Retirement Review*'
    }

).Count


Write-Host ""

Write-Host "==============================================" `
    -ForegroundColor Green

Write-Host " AUDIT COMPLETE" `
    -ForegroundColor Green

Write-Host "==============================================" `
    -ForegroundColor Green

Write-Host ""

Write-Host "Devices checked:    $($results.Count)"

Write-Host "Online:             $onlineTotal" `
    -ForegroundColor Green

Write-Host "Offline:            $offlineTotal" `
    -ForegroundColor Yellow

Write-Host "Owners unresolved:  $ownersUnresolved"

Write-Host "Retirement review:  $retirementTotal"

Write-Host ""

Write-Host "Excel report:" `
    -ForegroundColor Yellow

Write-Host $OutputWorkbook `
    -ForegroundColor Yellow

Write-Host ""

Write-Host "Technical CSV:" `
    -ForegroundColor DarkGray

Write-Host $checkpoint `
    -ForegroundColor DarkGray

Write-Host ""

Write-Host "Source CSV:" `
    -ForegroundColor DarkGray

Write-Host $InputCsv `
    -ForegroundColor DarkGray

Write-Host ""
