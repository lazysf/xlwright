<#
.SYNOPSIS
Stops a push that changes the plugin without raising its version.

.DESCRIPTION
Runs as a prek pre-push hook. prek sets PRE_COMMIT_FROM_REF (the commit the remote has) and
PRE_COMMIT_TO_REF (the commit being pushed).

Claude Code keeps users on the cached copy of a plugin until the version in plugin.json changes,
so a push whose range changes plugins/xlwright/ must also change that version. The version is read
from the commits themselves, not from the working tree.

When the remote does not have the branch yet, the push is compared with the merge base of the
remote's default branch (<remote>/HEAD, then <remote>/main). When neither exists, the check is
skipped with a warning. Deleting a remote ref is not checked.

Exit code 0 when the push may go ahead, 1 when it should stop.
#>
$ErrorActionPreference = 'Stop'

$pluginDir = 'plugins/xlwright/'
$manifest = 'plugins/xlwright/.claude-plugin/plugin.json'
$zero = '^0+$'

function Get-ManifestVersion([string]$Commit) {
    # ls-tree prints nothing when the file is missing, where show would fail.
    $listed = git ls-tree --name-only $Commit -- $manifest
    if ($LASTEXITCODE -ne 0) { throw "git ls-tree failed (exit $LASTEXITCODE)" }
    if (!$listed) { return $null }
    $json = (git show "${Commit}:$manifest") -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "git show failed (exit $LASTEXITCODE)" }
    $version = ($json | ConvertFrom-Json).version
    if ($version -is [string] -and $version) { return $version }
    return $null
}

function Resolve-Commit([string]$Ref) {
    $sha = git rev-parse --verify --quiet "$Ref^{commit}"
    if ($LASTEXITCODE -ne 0) { return $null }
    return $sha
}

$from = $env:PRE_COMMIT_FROM_REF
$to = $env:PRE_COMMIT_TO_REF
$remote = $env:PRE_COMMIT_REMOTE_NAME
if (!$remote) { $remote = 'origin' }

if (!$to) {
    Write-Output 'PRE_COMMIT_TO_REF is not set. Run this through prek at pre-push, or with'
    Write-Output '`prek run --hook-stage pre-push --from-ref <base> --to-ref <commit>`.'
    exit 1
}
if ($to -match $zero) {
    Write-Output 'Deleting a remote ref; nothing to check.'
    exit 0
}

# git writes the manifest as UTF-8; the console default (CP932) would garble it.
$savedEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
try {
    Push-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)
    try {
        $toCommit = Resolve-Commit $to
        if (!$toCommit) {
            Write-Output "Cannot find the commit being pushed ($to)."
            exit 1
        }

        $base = $null
        if ($from -and $from -notmatch $zero) { $base = Resolve-Commit $from }
        if (!$base) {
            foreach ($candidate in @("refs/remotes/$remote/HEAD", "refs/remotes/$remote/main")) {
                $tip = Resolve-Commit $candidate
                if (!$tip) { continue }
                $base = git merge-base $tip $toCommit
                if ($LASTEXITCODE -ne 0) { $base = $null; continue }
                Write-Output "The remote does not have this branch; comparing with the merge base of $candidate."
                break
            }
        }
        if (!$base) {
            # tradeoff: an unrelated history (or no remote-tracking ref) passes unchecked. Make this
            # stop the push if branches other than main start to carry releases.
            Write-Output "WARNING: found nothing to compare the push with; the plugin version was not checked."
            exit 0
        }

        $changed = @(git diff --name-only $base $toCommit -- $pluginDir)
        if ($LASTEXITCODE -ne 0) { throw "git diff failed (exit $LASTEXITCODE)" }
        if ($changed.Count -eq 0) {
            Write-Output "No changes under $pluginDir; the plugin version need not change."
            exit 0
        }

        $newVersion = Get-ManifestVersion $toCommit
        if (!$newVersion) {
            Write-Output "$manifest has no version in the commit being pushed."
            exit 1
        }
        $oldVersion = Get-ManifestVersion $base
        if ($oldVersion -eq $newVersion) {
            Write-Output "This push changes $($changed.Count) file(s) under $pluginDir but keeps version $newVersion."
            Write-Output "Users stay on the cached copy until the version changes. Raise `"version`" in $manifest."
            exit 1
        }
        Write-Output "Plugin version: $oldVersion -> $newVersion"
        exit 0
    }
    finally { Pop-Location }
}
finally {
    [Console]::OutputEncoding = $savedEncoding
}
