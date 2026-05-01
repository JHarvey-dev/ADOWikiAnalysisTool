---
title: ADO Wiki Stats
description: PowerShell script for exporting Azure DevOps wiki staleness and usage statistics to CSV.
---

## ADO Wiki Stats

A PowerShell tool that queries an **Azure DevOps (ADO)** wiki and exports per-page **staleness** and **usage** statistics to a CSV file.

---

## What it does

For every page in one or more ADO wikis the script collects:

| Column | Description |
|---|---|
| `WikiName` | Display name of the wiki |
| `WikiId` | Internal GUID of the wiki |
| `PagePath` | Full path of the page within the wiki (e.g. `/Engineering/Onboarding`) |
| `PageUrl` | Direct browser URL to the page |
| `LastUpdatedDate` | Timestamp of the most-recent git commit that changed this page |
| `LastUpdatedBy` | Author name from that commit |
| `PageViewsLast30Days` | Total page views over the configured look-back window (default 30 days) |

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ (Windows) **or** PowerShell 7+ (cross-platform) | Included with Windows; download [PowerShell 7](https://github.com/PowerShell/PowerShell/releases) for macOS/Linux |
| Azure DevOps Personal Access Token (PAT) | Must have **Wiki – Read** and **Code – Read** scopes |

### Creating a PAT

1. Sign in to `https://dev.azure.com/<YourOrganization>`.
2. Click your avatar (top-right) → **Personal access tokens** → **New Token**.
3. Give it a name, set an expiry, and tick:
   - **Wiki** → Read
   - **Code** → Read
4. Copy the generated token and store it safely. Avoid putting the token directly on the command line.

---

## Usage

```powershell
$env:ADO_PAT = "<your-pat>"

.\Get-ADOWikiStats.ps1 `
    -Organization "<org>" `
    -Project      "<project>"
```

This produces `WikiStats.csv` in the current directory containing stats for **all wikis** in the project.

### All parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-Organization` | ✅ | – | ADO organisation name (slug from the URL) |
| `-Project` | ✅ | – | ADO project name or GUID |
| `-PAT` | ❌ | Prompt/env | Personal Access Token. If omitted, script uses `ADO_PAT` or `AZURE_DEVOPS_EXT_PAT`, then prompts securely |
| `-WikiIdentifier` | ❌ | *(all wikis)* | Name or GUID of a single wiki to query |
| `-PageViewsForDays` | ❌ | `30` | Days of view history to include (1–30) |
| `-OutputFile` | ❌ | `WikiStats.csv` | Path of the output CSV file |

### Examples

**Query all wikis, last 30 days of stats:**
```powershell
$env:ADO_PAT = "<your-pat>"

.\Get-ADOWikiStats.ps1 `
    -Organization "contoso" `
    -Project      "PlatformTeam"
```

**Prompt for PAT securely (no env var and no `-PAT`):**
```powershell
.\Get-ADOWikiStats.ps1 `
    -Organization "contoso" `
    -Project      "PlatformTeam"
```

**Query a specific wiki with a custom output path:**
```powershell
.\Get-ADOWikiStats.ps1 `
    -Organization   "contoso" `
    -Project        "PlatformTeam" `
    -WikiIdentifier "PlatformTeam.wiki" `
    -OutputFile     "C:\Reports\PlatformWikiStats.csv"
```

**Query the last 14 days only:**
```powershell
.\Get-ADOWikiStats.ps1 `
    -Organization     "contoso" `
    -Project          "PlatformTeam" `
    -PageViewsForDays 14
```

**Enable verbose output to track progress per page:**
```powershell
.\Get-ADOWikiStats.ps1 `
    -Organization "contoso" `
    -Project      "PlatformTeam" `
    -Verbose
```

## Security

* Use `ADO_PAT` (or the secure prompt) instead of passing `-PAT` inline.
* Use a short PAT expiry with minimum scopes (`Wiki: Read`, `Code: Read`).
* If you find a security issue, report it privately to the maintainer.

---

## How it runs

The script runs in two phases:

1. **Discovery** — Fetches the full page tree for each wiki (one API call per wiki) and reports the total number of pages found.
2. **Prompt** — Asks you how to proceed:
   - **`A`** — Process all discovered pages
   - **A number** (e.g. `50`) — Process only the first N pages (useful for a quick test run)
   - **`Q`** — Quit without making further API calls
3. **Collection** — For each selected page, retrieves view stats and git commit history, showing a progress bar as it goes.

All API calls are read-only (`GET` requests). The script includes automatic retry with back-off if Azure DevOps returns HTTP 429 (rate-limited).

---

## How it works

```
Get-ADOWikiStats.ps1
│
├── 1. LIST WIKIS
│       GET /wiki/wikis
│       Enumerates all wikis in the project (or filters to the requested one).
│
├── 2. LIST PAGES  (one call per wiki)
│       GET /wiki/wikis/{id}/pages?recursionLevel=full
│       Fetches the full page tree and flattens it into a list.
│
├── 3. PAGE VIEW STATS  (one bulk call per wiki)
│       GET /wiki/wikis/{id}/pages/stats?pageViewsForDays=N
│       Returns view counts for every page in a single round-trip.
│
└── 4. LAST COMMIT  (one call per page)
        GET /git/repositories/{repoId}/commits?searchCriteria.itemPath={path}&$top=1
        Retrieves the most-recent git commit for each page file to determine
        when it was last edited and by whom.
```

The results from steps 2–4 are joined in memory and written to the CSV.

> **Performance note:** The script makes one git-commit API call per page.
> For very large wikis (hundreds of pages) this may take a few minutes.
> Use `-Verbose` to monitor progress.

---

## Output example

```
WikiName,WikiId,PagePath,PageId,LastUpdatedDate,LastUpdatedBy,PageViewsLast30Days,PageUrl
PlatformTeam.wiki,abc-123-...,/,1,2026-03-15T09:22:00Z,Alice Smith,45,https://dev.azure.com/...
PlatformTeam.wiki,abc-123-...,/Engineering,2,2026-01-04T14:05:00Z,Bob Jones,12,https://dev.azure.com/...
PlatformTeam.wiki,abc-123-...,/Engineering/Onboarding,3,2025-11-20T08:30:00Z,Alice Smith,120,https://dev.azure.com/...
```

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `401 Unauthorized` | PAT is invalid or expired | Regenerate the PAT and ensure it has Wiki + Code Read scopes |
| `No wikis found` | Project name is wrong, or no wiki exists | Verify the project name and that a wiki has been created |
| `Wiki 'X' not found` | `-WikiIdentifier` does not match any wiki name/GUID | Run without `-WikiIdentifier` first to see all available wikis |
| `LastUpdatedDate` is blank for some pages | Page exists in the wiki index but has no git commits (e.g. auto-generated root) | Expected – the page was never directly edited |
| Script runs slowly | Large wiki; one API call per page for git history | Normal behaviour – use `-Verbose` to monitor progress |
