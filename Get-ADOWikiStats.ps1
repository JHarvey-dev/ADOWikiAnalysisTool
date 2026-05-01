<#
.SYNOPSIS
    Queries an Azure DevOps wiki and exports page staleness and usage statistics to a CSV file.

.DESCRIPTION
    For every page in one or more Azure DevOps wikis this script collects:
      - Page path
      - Last-updated timestamp (derived from the most-recent git commit to that page file)
      - Author of the last update
      - Page view count over a configurable look-back window (default: 30 days)
      - Wiki name and wiki ID

    Results are written to a CSV file for easy review in Excel or Power BI.

.PARAMETER Organization
    Azure DevOps organisation name (the slug that appears in your ADO URL:
    https://dev.azure.com/<Organization>).

.PARAMETER Project
    Azure DevOps project name or GUID.

.PARAMETER PAT
        Optional. Personal Access Token with at least the following scopes:
      - Wiki (Read)
      - Code  (Read)   – required to query git commit history

        If omitted, the script tries environment variables first:
            - ADO_PAT
            - AZURE_DEVOPS_EXT_PAT
        If neither is set, the script prompts for the PAT securely.

.PARAMETER WikiIdentifier
    Optional. Name or GUID of a specific wiki to query.
    When omitted the script processes every wiki found in the project.

.PARAMETER PageViewsForDays
    Number of days of page-view history to retrieve (1–30, ADO maximum is 30).
    Defaults to 30.

.PARAMETER OutputFile
    Path of the CSV file to write. Defaults to "WikiStats.csv" in the script
    directory (or current directory if run interactively).

.PARAMETER ContinueFromPreviousRun
    When set, the script reads the existing output CSV to determine which pages
    have already been processed, skips them, and appends new rows to the file.
    Use this to process a large wiki in batches.

.EXAMPLE
    .\Get-ADOWikiStats.ps1 `
        -Organization "myorg" `
        -Project     "MyProject"

.EXAMPLE
    $env:ADO_PAT = "<your-pat>"

    .\Get-ADOWikiStats.ps1 `
        -Organization "myorg" `
        -Project     "MyProject"

.EXAMPLE
    .\Get-ADOWikiStats.ps1 `
        -Organization    "myorg" `
        -Project         "MyProject" `
        -WikiIdentifier  "MyProject.wiki" `
        -PageViewsForDays 14 `
        -OutputFile       "C:\Reports\WikiStats.csv"

.NOTES
    Requires PowerShell 5.1 or later (Windows) or PowerShell 7+ (cross-platform).
    ADO REST API version used: 7.1
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = 'ADO organisation name')]
    [ValidateNotNullOrEmpty()]
    [string]$Organization,

    [Parameter(Mandatory = $true, HelpMessage = 'ADO project name or GUID')]
    [ValidateNotNullOrEmpty()]
    [string]$Project,

    [Parameter(HelpMessage = 'Personal Access Token; omit to use env var or secure prompt')]
    [string]$PAT = '',

    [Parameter(HelpMessage = 'Wiki name or GUID; omit to query all wikis in the project')]
    [string]$WikiIdentifier = '',

    [Parameter(HelpMessage = 'Number of days of page-view history (1-30)')]
    [ValidateRange(1, 30)]
    [int]$PageViewsForDays = 30,

    [Parameter(HelpMessage = 'Output CSV file path')]
    [string]$OutputFile = '',

    [Parameter(HelpMessage = 'Resume from a previous run; appends to existing CSV')]
    [switch]$ContinueFromPreviousRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $OutputFile) {
    $root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
    $OutputFile = Join-Path $root 'WikiStats.csv'
}

if (-not $PAT) {
    if ($env:ADO_PAT) {
        $PAT = $env:ADO_PAT
    }
    elseif ($env:AZURE_DEVOPS_EXT_PAT) {
        $PAT = $env:AZURE_DEVOPS_EXT_PAT
    }
    else {
        $securePat = Read-Host -Prompt 'Enter Azure DevOps PAT' -AsSecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePat)
        try {
            $PAT = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

if (-not $PAT) {
    Write-Error "No PAT was provided. Pass -PAT, set ADO_PAT/AZURE_DEVOPS_EXT_PAT, or enter a PAT when prompted."
    exit 1
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function New-AuthHeader {
    param ([string]$Pat)
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    return @{ Authorization = "Basic $encoded" }
}

function Invoke-AdoApi {
    <#
    .SYNOPSIS Thin wrapper around Invoke-RestMethod with retry on throttle (429).
    #>
    param (
        [string]$Uri,
        [hashtable]$Headers
    )

    $maxRetries = 10
    $attempt    = 0

    while ($true) {
        $attempt++
        try {
            $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -ContentType 'application/json'

            # ADO can sometimes return an HTML sign-in page instead of JSON when auth fails.
            if ($response -is [string] -and $response -match '<html|<HTML|Sign in|signin') {
                throw "Authentication failed while calling '$Uri'. Received HTML instead of JSON. Verify PAT validity and scopes."
            }

            return $response
        }
        catch {
            $statusCode = $null
            $retryAfter = $null
            if ($_.Exception.PSObject.Properties.Match('Response').Count -gt 0 -and $_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                # Try to read Retry-After header from ADO
                try {
                    $raHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -ieq 'Retry-After' }
                    if ($raHeader) { $retryAfter = [int]$raHeader.Value[0] }
                } catch { }
            }
            if ($statusCode -in @(401, 403)) {
                throw "Authentication/authorization failed (HTTP $statusCode) calling '$Uri'. Verify PAT validity and scopes: Wiki (Read), Code (Read), and project access."
            }
            if ($statusCode -eq 429 -and $attempt -le $maxRetries) {
                if (-not $retryAfter -or $retryAfter -lt 1) { $retryAfter = 10 * $attempt }
                Write-Warning "Rate-limited (429). Waiting $retryAfter s before retry $attempt/$maxRetries..."
                Start-Sleep -Seconds $retryAfter
            }
            else {
                throw
            }
        }
    }
}

function Get-PropertyValueOrDefault {
    param (
        [object]$InputObject,
        [string]$PropertyName,
        [object]$DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function Get-AllWikiPages {
    <#
    .SYNOPSIS
        Returns a flat list of all WikiPage objects from the given wiki using
        a single GET with recursionLevel=full, then flattening the nested tree.
    #>
    param (
        [string]$BaseUrl,
        [string]$WikiId,
        [hashtable]$Headers
    )

    $uri = "$BaseUrl/_apis/wiki/wikis/$WikiId/pages?recursionLevel=full&includeContent=false&api-version=7.1"
    Write-Verbose "  Fetching full page tree: $uri"

    try {
        $root = Invoke-AdoApi -Uri $uri -Headers $Headers
    }
    catch {
        Write-Warning "  Could not retrieve pages for wiki '$WikiId': $_"
        return @()
    }

    # Flatten the nested subPages tree into a flat list using a queue
    $pages = [System.Collections.Generic.List[object]]::new()
    $queue  = [System.Collections.Generic.Queue[object]]::new()
    $queue.Enqueue($root)

    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        # The root '/' has no gitItemPath — skip it
        if ($node.PSObject.Properties.Match('gitItemPath').Count -gt 0 -and $node.gitItemPath) {
            $pages.Add($node)
        }
        if ($node.PSObject.Properties.Match('subPages').Count -gt 0 -and $node.subPages) {
            foreach ($child in $node.subPages) {
                $queue.Enqueue($child)
            }
        }
    }

    return $pages
}

function Get-WikiPageStats {
    <#
    .SYNOPSIS
        Returns the total page view count for a single wiki page using its path.
    #>
    param (
        [string]$BaseUrl,
        [string]$WikiId,
        [string]$PagePath,
        [int]$Days,
        [hashtable]$Headers
    )

    $encodedPath = [Uri]::EscapeDataString($PagePath)
    $uri = "$BaseUrl/_apis/wiki/wikis/$WikiId/pages?path=$encodedPath&recursionLevel=none&includeContent=false&api-version=7.1"
    Write-Verbose "  Fetching page id for: $PagePath"

    try {
        $pageDetail = Invoke-AdoApi -Uri $uri -Headers $Headers
        $pageId = $pageDetail.id
    }
    catch {
        Write-Warning "  Could not retrieve page details for '$PagePath' in wiki '$WikiId': $_"
        return 0
    }

    $statsUri = "$BaseUrl/_apis/wiki/wikis/$WikiId/pages/$pageId/stats?pageViewsForDays=$Days&api-version=7.1"
    Write-Verbose "  Fetching page stats: $statsUri"

    try {
        $result = Invoke-AdoApi -Uri $statsUri -Headers $Headers
        $totalViews = 0
        if ($result.PSObject.Properties.Match('viewStats').Count -gt 0 -and $result.viewStats) {
            foreach ($day in $result.viewStats) {
                $totalViews += [int]$day.count
            }
        }
        return $totalViews
    }
    catch {
        Write-Warning "  Could not retrieve page stats for '$PagePath' in wiki '$WikiId': $_"
        return 0
    }
}

function Get-PageLastCommit {
    <#
    .SYNOPSIS
        Returns a PSCustomObject with AuthorDate and AuthorName for the most
        recent git commit that touched the given wiki page file.
    #>
    param (
        [string]$BaseUrl,
        [string]$GitRepoId,
        [string]$GitItemPath,
        [hashtable]$Headers
    )

    if (-not $GitItemPath) {
        return [PSCustomObject]@{ AuthorDate = $null; AuthorName = '' }
    }

    $encodedPath = [Uri]::EscapeDataString($GitItemPath)
    $uri = "$BaseUrl/_apis/git/repositories/$GitRepoId/commits" +
           "?searchCriteria.itemPath=$encodedPath&`$top=1&api-version=7.1"

    try {
        $result = Invoke-AdoApi -Uri $uri -Headers $Headers
        if ($result.value -and @($result.value).Count -gt 0) {
            $commit = $result.value[0]
            return [PSCustomObject]@{
                AuthorDate = $commit.author.date
                AuthorName = $commit.author.name
            }
        }
    }
    catch {
        Write-Warning "    Could not get git commits for '$GitItemPath': $_"
    }

    return [PSCustomObject]@{ AuthorDate = $null; AuthorName = '' }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$headers = New-AuthHeader -Pat $PAT
$baseUrl  = "https://dev.azure.com/$Organization/$([Uri]::EscapeDataString($Project))"

# Column name is derived from the actual look-back period so the CSV header
# is accurate regardless of whether the default (30) or a custom value is used.
$viewsColumnName = "PageViewsLast${PageViewsForDays}Days"

# 1. Discover wikis
Write-Host "Fetching wiki list from '$Organization/$Project'..."
$wikisUri = "$baseUrl/_apis/wiki/wikis?api-version=7.1"
$wikisResponse = Invoke-AdoApi -Uri $wikisUri -Headers $headers

if ($wikisResponse -is [string] -or $wikisResponse.PSObject.Properties.Match('value').Count -eq 0) {
    $serverMessage = Get-PropertyValueOrDefault -InputObject $wikisResponse -PropertyName 'message' -DefaultValue ''
    Write-Error "Unexpected wiki-list response shape. This is commonly caused by an invalid/expired PAT or missing access to '$Project'. $serverMessage"
    exit 1
}

$wikis = @($wikisResponse.value)
if ($wikis.Count -eq 0) {
    Write-Error "No wikis found in project '$Project'."
    exit 1
}

if ($WikiIdentifier) {
    $wikis = @($wikis | Where-Object { $_.id -eq $WikiIdentifier -or $_.name -eq $WikiIdentifier })
    if (-not $wikis) {
        Write-Error "Wiki '$WikiIdentifier' not found in project '$Project'."
        exit 1
    }
}

Write-Host "Found $($wikis.Count) wiki(s) to process."

# ---------------------------------------------------------------------------
# Phase 1 – Discover all pages across all wikis
# ---------------------------------------------------------------------------
$wikiPages = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($wiki in $wikis) {
    Write-Host ""
    Write-Host "Discovering pages in wiki: '$($wiki.name)' ..."
    $pages = @(Get-AllWikiPages -BaseUrl $baseUrl -WikiId $wiki.id -Headers $headers)
    Write-Host "  $($pages.Count) page(s) found."
    foreach ($p in $pages) {
        $wikiPages.Add([PSCustomObject]@{
            Wiki = $wiki
            Page = $p
        })
    }
}

$totalPages = $wikiPages.Count
if ($totalPages -eq 0) {
    Write-Host "No pages found in any wiki. Nothing to do."
    exit 0
}

# --- Resume support: read existing CSV to find already-processed pages ------
$alreadyProcessed = @{}
if ($ContinueFromPreviousRun -and (Test-Path $OutputFile)) {
    $existing = Import-Csv -Path $OutputFile
    foreach ($row in $existing) {
        $key = "$($row.WikiId)|$($row.PagePath)"
        $alreadyProcessed[$key] = $true
    }
    Write-Host "  Loaded $($alreadyProcessed.Count) already-processed page(s) from previous run."
}

# Filter out already-processed pages
if ($alreadyProcessed.Count -gt 0) {
    $remaining = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($entry in $wikiPages) {
        $key = "$($entry.Wiki.id)|$($entry.Page.path)"
        if (-not $alreadyProcessed.ContainsKey($key)) {
            $remaining.Add($entry)
        }
    }
    $wikiPages = $remaining
    Write-Host "  $($wikiPages.Count) page(s) remaining after skipping previous results."
    if ($wikiPages.Count -eq 0) {
        Write-Host "All pages already processed. Nothing to do."
        exit 0
    }
}

$remainingCount = $wikiPages.Count
$apiCalls = $remainingCount * 3
Write-Host ""
Write-Host "=========================================="
Write-Host "  Total pages discovered: $totalPages"
if ($alreadyProcessed.Count -gt 0) {
    Write-Host "  Already processed:      $($alreadyProcessed.Count)"
}
Write-Host "  Remaining to process:   $remainingCount"
Write-Host "  Estimated API calls:    ~$apiCalls"
Write-Host "=========================================="
Write-Host ""
Write-Host "Options:"
Write-Host "  [A] Process ALL $remainingCount remaining pages"
Write-Host "  [N] Process next N pages (enter a number)"
Write-Host "  [Q] Quit"
Write-Host ""

$choice = Read-Host "Your choice"

switch -Regex ($choice.Trim()) {
    '^[Aa]$' {
        $limit = $remainingCount
    }
    '^[Qq]$' {
        Write-Host "Exiting."
        exit 0
    }
    '^\d+$' {
        $limit = [math]::Min([int]$choice, $remainingCount)
        if ($limit -le 0) {
            Write-Host "Invalid number. Exiting."
            exit 1
        }
        Write-Host "Processing next $limit of $remainingCount remaining pages."
    }
    default {
        Write-Host "Unrecognized input '$choice'. Exiting."
        exit 1
    }
}

# ---------------------------------------------------------------------------
# Phase 2 – Collect stats and git history for selected pages
# ---------------------------------------------------------------------------

# Prepare CSV file — write header if starting fresh, or verify append mode
$isAppend = $ContinueFromPreviousRun -and (Test-Path $OutputFile) -and $alreadyProcessed.Count -gt 0
if (-not $isAppend) {
    # Write header row
    $headerRow = '"WikiName","WikiId","PagePath","PageUrl","LastUpdatedDate","LastUpdatedBy","{0}"' -f $viewsColumnName
    [System.IO.File]::WriteAllText($OutputFile, "$headerRow`r`n", [System.Text.UTF8Encoding]::new($false))
}

$index       = 0
$rowsWritten = 0

foreach ($entry in $wikiPages) {
    if ($index -ge $limit) { break }
    $index++

    $wiki = $entry.Wiki
    $page = $entry.Page

    Write-Progress -Activity "Collecting stats" `
        -Status "$index of $limit pages" `
        -PercentComplete ([math]::Floor(($index / $limit) * 100)) `
        -CurrentOperation $page.path

    $lastCommit = Get-PageLastCommit `
        -BaseUrl     $baseUrl `
        -GitRepoId   $wiki.repositoryId `
        -GitItemPath $page.gitItemPath `
        -Headers     $headers

    $totalViews = Get-WikiPageStats `
        -BaseUrl  $baseUrl `
        -WikiId   $wiki.id `
        -PagePath $page.path `
        -Days     $PageViewsForDays `
        -Headers  $headers

    $row = [PSCustomObject][ordered]@{
        WikiName        = $wiki.name
        WikiId          = $wiki.id
        PagePath        = $page.path
        PageUrl         = $page.remoteUrl
        LastUpdatedDate = if ($lastCommit.AuthorDate) { [datetime]$lastCommit.AuthorDate } else { $null }
        LastUpdatedBy   = $lastCommit.AuthorName
        $viewsColumnName = $totalViews
    }

    # Append this row immediately so progress is never lost
    $csvLine = ($row | ConvertTo-Csv -NoTypeInformation | Select-Object -Last 1)
    [System.IO.File]::AppendAllText($OutputFile, "$csvLine`r`n", [System.Text.UTF8Encoding]::new($false))
    $rowsWritten++
}
Write-Progress -Activity "Collecting stats" -Completed

$totalInFile = $rowsWritten + $alreadyProcessed.Count
$pagesLeft   = $totalPages - $totalInFile
Write-Host ""
Write-Host "Wrote $rowsWritten row(s) this run. Total in file: $totalInFile of $totalPages."
if ($pagesLeft -gt 0) {
    Write-Host "$pagesLeft page(s) remaining. Re-run with -ContinueFromPreviousRun to process the next batch."
}
Write-Host "Output: $(Resolve-Path $OutputFile)"
