param(
    [switch]$Force
)

$cmdPath = Join-Path $PSScriptRoot "convert-to-exec-deck.cmd"
if (-not (Test-Path $cmdPath)) {
    Write-Host "convert-to-exec-deck.cmd not found in repo root." -ForegroundColor Red
    exit 1
}

$profilePath = $PROFILE
$markerStart = "# BEGIN pbi-to-exec-deck alias"
$markerEnd = "# END pbi-to-exec-deck alias"

$aliasBlock = @"
$markerStart
function Convert-To-Exec-Deck {
    & `"$cmdPath`" @Args
}
Set-Alias convert-to-exec-deck Convert-To-Exec-Deck
$markerEnd
"@

if (-not (Test-Path $profilePath)) {
    New-Item -ItemType File -Path $profilePath -Force | Out-Null
}

$profileContent = Get-Content -Path $profilePath -Raw
if ($profileContent -match [regex]::Escape($markerStart)) {
    if (-not $Force) {
        Write-Host "Alias already installed in profile. Use -Force to reinstall." -ForegroundColor Yellow
        exit 0
    }
    $pattern = [regex]::Escape($markerStart) + ".*?" + [regex]::Escape($markerEnd)
    $profileContent = [regex]::Replace($profileContent, $pattern, $aliasBlock, "Singleline")
    Set-Content -Path $profilePath -Value $profileContent
} else {
    Add-Content -Path $profilePath -Value ("`n" + $aliasBlock)
}

Write-Host "Alias installed. Reload your profile or open a new PowerShell window." -ForegroundColor Green
