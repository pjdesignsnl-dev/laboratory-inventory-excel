[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://github\.com/[^/]+/[^/]+(?:\.git)?$')]
    [string]$RemoteUrl
)

$ErrorActionPreference = 'Stop'

function Invoke-Git {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Arguments -join ' ') failed with exit code $LASTEXITCODE."
    }
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw 'Git is not installed or is not available on PATH.'
}

$repoRoot = (& git rev-parse --show-toplevel 2>$null)
if ($LASTEXITCODE -ne 0 -or -not $repoRoot) {
    throw 'Run this script from inside the cloned/extracted Git repository.'
}

Set-Location $repoRoot

$currentBranch = (& git branch --show-current).Trim()
if ($LASTEXITCODE -ne 0 -or $currentBranch -ne 'main') {
    throw "Expected current branch 'main'; found '$currentBranch'."
}

$dirty = & git status --porcelain
if ($LASTEXITCODE -ne 0) {
    throw 'Could not inspect Git status.'
}
if ($dirty) {
    throw 'Working tree is not clean. Review or commit local changes before publishing.'
}

$existingOrigin = & git remote get-url origin 2>$null
if ($LASTEXITCODE -eq 0 -and $existingOrigin) {
    $existingOrigin = $existingOrigin.Trim()
    $requestedOrigin = $RemoteUrl.Trim()

    if ($existingOrigin -eq $requestedOrigin) {
        Write-Host "Origin already configured: $existingOrigin"
    }
    elseif ($existingOrigin -match '\.bundle$' -or (Test-Path -LiteralPath $existingOrigin)) {
        # A repository cloned from a Git bundle receives the local bundle path as
        # its origin. Replace only that clearly local bootstrap origin.
        Write-Host "Replacing local bundle origin: $existingOrigin"
        Invoke-Git remote set-url origin $requestedOrigin
    }
    else {
        throw "An origin remote already exists with a different non-local URL: $existingOrigin"
    }
}
else {
    Invoke-Git remote add origin $RemoteUrl
}

Write-Host 'Pushing main branch...'
Invoke-Git push -u origin main

Write-Host 'Pushing tags...'
Invoke-Git push origin --tags

Write-Host ''
Write-Host 'Published successfully.'
Invoke-Git remote -v
Invoke-Git status --short --branch
Invoke-Git log -1 --oneline
