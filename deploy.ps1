# amsong72 deploy: sync gas/index.html -> syntax check -> clasp push/version/redeploy (same web app URL) -> git commit/push
# Usage (repo root):  .\deploy.ps1 -Msg "change description"
# Windows PowerShell 5.1 compatible. Any failed step exits with code 1.
param([Parameter(Mandatory=$true)][string]$Msg)
$ErrorActionPreference = 'Continue'

$DeployId = 'AKfycbwCbl2SnhEF1fj-zf35ptta4p3bqzP1YLrb4OQdVk8GTlFtR_w7ewfGu5WZY2z1bWyt'
$ExecUrl = "https://script.google.com/macros/s/$DeployId/exec"
$Root = $PSScriptRoot
$GasDir = Join-Path $Root 'gas'
$script:Pushed = $false

function Fail([string]$m) {
  if ($script:Pushed) { Pop-Location; $script:Pushed = $false }
  Write-Host "[deploy] FAILED: $m" -ForegroundColor Red
  exit 1
}
function Step([string]$m) { Write-Host "[deploy] $m" -ForegroundColor Cyan }

# (1) copy root index.html -> gas/index.html and verify identical
Step '1/5 sync gas/index.html'
$src = Join-Path $Root 'index.html'
$dst = Join-Path $GasDir 'index.html'
if (-not (Test-Path $src)) { Fail "missing $src" }
try { Copy-Item $src $dst -Force -ErrorAction Stop } catch { Fail "copy failed: $_" }
$h1 = (Get-FileHash $src -Algorithm SHA256).Hash
$h2 = (Get-FileHash $dst -Algorithm SHA256).Hash
if ($h1 -ne $h2) { Fail 'index.html and gas/index.html differ after copy' }

# (2) extract <script>...</script> from index.html and run node --check
Step '2/5 syntax check (node --check)'
$node = Get-Command node -ErrorAction SilentlyContinue
if ($null -eq $node) {
  Write-Host '[deploy] WARNING: node not found - skipping syntax check' -ForegroundColor Yellow
} else {
  $html = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)
  $ms = [regex]::Matches($html, '(?s)<script[^>]*>(.*?)</script>')
  if ($ms.Count -eq 0) { Fail 'no <script> block found in index.html' }
  $parts = @()
  foreach ($m in $ms) { $parts += $m.Groups[1].Value }
  $tmp = Join-Path $env:TEMP ("amsong72_check_" + $PID + ".js")
  [System.IO.File]::WriteAllText($tmp, ($parts -join "`n;`n"), (New-Object System.Text.UTF8Encoding $false))
  & node --check $tmp
  $code = $LASTEXITCODE
  Remove-Item $tmp -Force -ErrorAction SilentlyContinue
  if ($code -ne 0) { Fail 'node --check failed for index.html script' }
}

# (3) clasp push -> version -> redeploy (keeps the same /exec URL)
Step '3/5 clasp push / version / redeploy'
Push-Location $GasDir
$script:Pushed = $true
& clasp push -f
if ($LASTEXITCODE -ne 0) { Fail 'clasp push failed' }
$out = & clasp version "$Msg" 2>&1 | Out-String
Write-Host $out
$vm = [regex]::Match($out, 'Created version (\d+)')
if (-not $vm.Success) { Fail 'could not parse version number from clasp version output' }
$n = $vm.Groups[1].Value
& clasp redeploy $DeployId -V $n -d "$Msg"
if ($LASTEXITCODE -ne 0) { Fail "clasp redeploy failed (version $n was created)" }
Pop-Location
$script:Pushed = $false

# (4) git commit + push (also updates GitHub Pages)
Step '4/5 git commit / push'
& git -C $Root add -A
if ($LASTEXITCODE -ne 0) { Fail 'git add failed' }
$pending = & git -C $Root status --porcelain
if ($pending) {
  & git -C $Root commit -m "$Msg" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
  if ($LASTEXITCODE -ne 0) { Fail 'git commit failed' }
} else {
  Write-Host '[deploy] nothing to commit'
}
& git -C $Root push
if ($LASTEXITCODE -ne 0) { Fail 'git push failed' }

# (5) result
Step '5/5 done'
Write-Host "[deploy] version: $n" -ForegroundColor Green
Write-Host "[deploy] web app: $ExecUrl" -ForegroundColor Green
exit 0
