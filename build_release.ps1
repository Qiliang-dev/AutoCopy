# Usage:
#   powershell -ExecutionPolicy Bypass -File .\build_release.ps1 -Version 1.8 -Name "AutoCopyTool_v1.8"
# Only required param is -Version; -Name defaults to "AutoCopyTool_v<Version>"

param(
    [Parameter(Mandatory=$true)]
    [string]$Version,

    [string]$Name
)

if (-not $Name -or $Name.Trim() -eq "") {
    $Name = "AutoCopyTool_v$Version"
}

Write-Host "Building $Name ..."

# Ensure venv not required; use system python
py -V *> $null 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "Python launcher (py) not found. Please install Python or add to PATH."
    exit 1
}

# Clean previous build artifacts for this version
Remove-Item -Recurse -Force "build" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "dist" -ErrorAction SilentlyContinue
Remove-Item "$Name.spec" -ErrorAction SilentlyContinue

# Build (embed icon if exists)
$iconPath = Join-Path "resources/icons" "autocopy.ico"
if (Test-Path $iconPath) {
    python -m PyInstaller --onefile --noconsole --name="$Name" --icon="$iconPath" auto_copy_gui.py --clean --noconfirm | cat
} else {
    python -m PyInstaller --onefile --noconsole --name="$Name" auto_copy_gui.py --clean --noconfirm | cat
}
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# Prepare release folder
$releaseDir = Join-Path "releases" "${Name}_Release"
New-Item -ItemType Directory -Force $releaseDir | Out-Null

Copy-Item "dist\$Name.exe" $releaseDir
Copy-Item "auto_copy_gui.py" $releaseDir
Copy-Item "requirements.txt" $releaseDir -ErrorAction SilentlyContinue

# Minimal README
$readme = @()
$readme += "AutoCopy Tool $Version"
$readme += "Build Name: $Name"
$readme += "Built at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$readme += ""
$readme += "- Changes: See commit log or release notes"
Set-Content -Path (Join-Path $releaseDir "README_${Version}.txt") -Value ($readme -join "`n")

# Zip
$zipPath = Join-Path "releases" "${Name}_Release.zip"
Compress-Archive -Path "$releaseDir\*" -DestinationPath $zipPath -Force

Write-Host "Done. Release at: $releaseDir and $zipPath"

