param(
  [string]$ProjectDir = (Get-Location).Path
)

$ErrorActionPreference = "Stop"

$GitCmd = "C:\Program Files\Git\cmd"
if ((Test-Path $GitCmd) -and (($env:Path -split ';') -notcontains $GitCmd)) {
  $env:Path = "$GitCmd;$env:Path"
}

$LogDir = Join-Path $env:USERPROFILE ".codex\logs"
$LogFile = Join-Path $LogDir "fund-investment-autosync.log"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Write-SyncLog {
  param([string]$Message)
  "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message" | Add-Content -Path $LogFile
}

function Invoke-GitQuiet {
  param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$GitArgs
  )

  $PreviousErrorActionPreference = $ErrorActionPreference
  $ErrorActionPreference = "Continue"

  try {
    git @GitArgs *> $null
    if ($LASTEXITCODE -ne 0) {
      throw "git $($GitArgs -join ' ') failed with exit code $LASTEXITCODE"
    }
  } finally {
    $ErrorActionPreference = $PreviousErrorActionPreference
  }
}

Set-Location $ProjectDir

try {
  Invoke-GitQuiet rev-parse --is-inside-work-tree
} catch {
  Write-SyncLog "Not a git repository."
  exit 0
}

try {
  Invoke-GitQuiet remote get-url origin
} catch {
  Write-SyncLog "No origin remote configured; skipping sync."
  exit 0
}

$Branch = (git branch --show-current).Trim()
if ([string]::IsNullOrWhiteSpace($Branch)) {
  Write-SyncLog "No current branch; skipping sync."
  exit 0
}

$Status = git status --porcelain
if ($Status) {
  Invoke-GitQuiet add -A
  $Staged = git diff --cached --name-only
  if ($Staged) {
    Invoke-GitQuiet commit -m "autosync: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-SyncLog "Committed local changes."
  }
}

try {
  Invoke-GitQuiet fetch origin $Branch
  Invoke-GitQuiet rev-parse --verify "origin/$Branch"
  Invoke-GitQuiet pull --rebase --autostash origin $Branch
} catch {
  Write-SyncLog "Remote branch origin/$Branch not found or pull failed; will try push."
}

try {
  Invoke-GitQuiet push -u origin $Branch
  Write-SyncLog "Pushed $Branch."
} catch {
  Write-SyncLog "Push failed."
}
