#Requires -Version 5.1
# ============================================================
# MinervaDB XL View - Windows Automated Installer
# install/install_windows.ps1
# ============================================================
# Installs all dependencies for MinervaDB XL View on Windows
# Requires: Windows 10/11 (64-bit), PowerShell 5.1+
#
# Usage (run as Administrator):
#   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
#   .\install\install_windows.ps1
# ============================================================

param(
    [switch]$SkipPython,
    [switch]$SkipODBC,
    [switch]$SkipGit,
    [string]$DSNName = "MinervaDB_XL_View",
    [string]$PGHost = "",
    [string]$PGPort = "5432",
    [string]$PGDatabase = "",
    [string]$PGUser = "excel_dashboard_ro"
)

$ErrorActionPreference = "Stop"

# ---- Helper Functions ----
function Write-Info    { param($msg) Write-Host "[INFO]  $msg" -ForegroundColor Cyan }
function Write-Success { param($msg) Write-Host "[OK]    $msg" -ForegroundColor Green }
function Write-Warn    { param($msg) Write-Host "[WARN]  $msg" -ForegroundColor Yellow }
function Write-Fail    { param($msg) Write-Host "[ERROR] $msg" -ForegroundColor Red; exit 1 }
function Write-Section { param($msg) Write-Host "`n==> $msg" -ForegroundColor Blue -BackgroundColor Black }

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir

Write-Host ""
Write-Host "  +----------------------------------------------+" -ForegroundColor Blue
Write-Host "  |   MinervaDB XL View Installer                |" -ForegroundColor Blue
Write-Host "  |   Windows Setup Script v1.0                  |" -ForegroundColor Blue
Write-Host "  +----------------------------------------------+" -ForegroundColor Blue
Write-Host ""

# ============================================================
# Step 1: Check Windows Version and Architecture
# ============================================================
Write-Section "Step 1: System Check"

$OSInfo = Get-CimInstance Win32_OperatingSystem
$OSArch = $OSInfo.OSArchitecture
$OSVer  = $OSInfo.Caption

Write-Info "OS: $OSVer ($OSArch)"

if ($OSArch -notmatch "64") {
    Write-Fail "MinervaDB XL View requires a 64-bit version of Windows."
}
Write-Success "64-bit Windows confirmed"

# Check PowerShell version
Write-Info "PowerShell version: $($PSVersionTable.PSVersion)"

# ============================================================
# Step 2: Install Git (if not present)
# ============================================================
Write-Section "Step 2: Git"

if ($SkipGit) {
    Write-Warn "Skipping Git installation (--SkipGit)"
} elseif (Get-Command git -ErrorAction SilentlyContinue) {
    Write-Success "Git already installed: $(git --version)"
} else {
    Write-Info "Installing Git via winget..."
    try {
        winget install --id Git.Git -e --source winget --silent --accept-package-agreements --accept-source-agreements
        $env:PATH += ";C:\Program Files\Git\cmd"
        Write-Success "Git installed"
    } catch {
        Write-Warn "winget failed. Please install Git manually from https://git-scm.com/download/win"
    }
}

# ============================================================
# Step 3: Install Python 3.12
# ============================================================
Write-Section "Step 3: Python"

if ($SkipPython) {
    Write-Warn "Skipping Python installation (--SkipPython)"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $PyVer = python --version 2>&1
    Write-Success "Python already installed: $PyVer"
} else {
    Write-Info "Installing Python 3.12 via winget..."
    try {
        winget install --id Python.Python.3.12 -e --source winget --silent --accept-package-agreements --accept-source-agreements
        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        Write-Success "Python 3.12 installed"
    } catch {
        Write-Warn "winget failed. Download Python from https://www.python.org/downloads/windows/"
    }
}

# Verify pip
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Warn "pip not found. Running: python -m ensurepip"
    python -m ensurepip --upgrade
}

# ============================================================
# Step 4: Install PostgreSQL ODBC Driver
# ============================================================
Write-Section "Step 4: PostgreSQL ODBC Driver (psqlODBC)"

if ($SkipODBC) {
    Write-Warn "Skipping ODBC installation (--SkipODBC)"
} else {
    # Check if psqlODBC is already installed
    $ODBCDrivers = Get-OdbcDriver -Name "PostgreSQL*" -ErrorAction SilentlyContinue
    if ($ODBCDrivers) {
        Write-Success "PostgreSQL ODBC driver already installed: $($ODBCDrivers[0].Name)"
    } else {
        Write-Info "Attempting to install psqlODBC via winget..."
        try {
            winget install --id PostgreSQL.psqlODBC -e --source winget --silent --accept-package-agreements --accept-source-agreements 2>$null
            Write-Success "psqlODBC installed via winget"
        } catch {
            Write-Warn "Automatic ODBC install failed."
            Write-Warn "Please download psqlODBC manually from:"
            Write-Warn "  https://www.postgresql.org/ftp/odbc/versions/msi/"
            Write-Warn "Download: psqlodbc_xx_xx_xxxx-x64.zip (64-bit)"
        }
    }
}

# ============================================================
# Step 5: Create Python Virtual Environment
# ============================================================
Write-Section "Step 5: Python Virtual Environment"

Set-Location $RepoRoot

if (Test-Path ".venv") {
    Write-Warn "Virtual environment already exists, skipping creation"
} else {
    Write-Info "Creating virtual environment..."
    python -m venv .venv
    Write-Success "Virtual environment created at .venv\"
}

# Activate
.venv\Scripts\Activate.ps1
Write-Info "Virtual environment activated"

# ============================================================
# Step 6: Install Python Dependencies
# ============================================================
Write-Section "Step 6: Python Dependencies"

Write-Info "Upgrading pip..."
python -m pip install --upgrade pip --quiet

Write-Info "Installing MinervaDB XL View dependencies..."
pip install -r python\requirements.txt

# Verify packages
$packages = @("psycopg2", "pandas", "openpyxl", "sqlalchemy")
foreach ($pkg in $packages) {
    $ver = python -c "import $pkg; print($pkg.__version__)" 2>$null
    if ($ver) {
        Write-Success "$pkg $ver"
    } else {
        Write-Warn "$pkg not found, trying binary..."
        pip install "${pkg}-binary" 2>$null
        if ($LASTEXITCODE -ne 0) { pip install $pkg }
    }
}

# ============================================================
# Step 7: Configure MinervaDB XL View
# ============================================================
Write-Section "Step 7: Configuration"

if (Test-Path "config\config.ini") {
    Write-Warn "config\config.ini already exists, skipping"
} else {
    Copy-Item "config\config.example.ini" "config\config.ini"
    Write-Success "Created config\config.ini from template"
    Write-Warn "Edit config\config.ini with your PostgreSQL credentials"
}

# ============================================================
# Step 8: Create ODBC DSN
# ============================================================
Write-Section "Step 8: ODBC Data Source (DSN)"

$ODBCDrivers = Get-OdbcDriver -Name "PostgreSQL Unicode*" -ErrorAction SilentlyContinue

if (-not $ODBCDrivers) {
    Write-Warn "PostgreSQL Unicode ODBC driver not found. Skipping DSN creation."
    Write-Warn "Install psqlODBC first, then re-run this script."
} elseif ($PGHost -eq "") {
    Write-Warn "No --PGHost provided. Skipping automatic DSN creation."
    Write-Info "To create DSN manually:"
    Write-Info "  1. Run odbcad32.exe"
    Write-Info "  2. User DSN -> Add -> PostgreSQL Unicode"
    Write-Info "  3. Name: $DSNName"
} else {
    Write-Info "Creating ODBC DSN: $DSNName..."
    $DSNParams = @{
        DsnName    = $DSNName
        DsnType    = "User"
        DriverName = "PostgreSQL Unicode"
        SetPropertyValue = @(
            "Servername=$PGHost",
            "Port=$PGPort",
            "Database=$PGDatabase",
            "Username=$PGUser"
        )
    }
    try {
        Add-OdbcDsn @DSNParams
        Write-Success "ODBC DSN created: $DSNName"
    } catch {
        Write-Warn "Could not create DSN automatically: $_"
        Write-Warn "Create it manually via odbcad32.exe"
    }
}

# ============================================================
# Step 9: Create output directory
# ============================================================
Write-Section "Step 9: Output Directory"
New-Item -ItemType Directory -Path "output" -Force | Out-Null
Write-Success "Output directory ready at .\output\"

# ============================================================
# Step 10: Test Connection
# ============================================================
Write-Section "Step 10: Connection Test"

$ConfigHost = (Select-String -Path "config\config.ini" -Pattern "^host\s*=" | ForEach-Object { $_.Line -replace ".*=\s*", "" }).Trim()

if ($ConfigHost -eq "" -or $ConfigHost -eq "your-postgresql-host") {
    Write-Warn "PostgreSQL host not configured. Edit config\config.ini and then run:"
    Write-Info "  .venv\Scripts\Activate.ps1"
    Write-Info "  python python\pg_connector.py"
} else {
    Write-Info "Testing PostgreSQL connection..."
    $TestResult = python python\pg_connector.py 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Connection test passed!"
    } else {
        Write-Warn "Connection test failed. Check config\config.ini settings."
        Write-Warn $TestResult
    }
}

# ============================================================
# Install Complete
# ============================================================
Write-Host ""
Write-Host "  +----------------------------------------------+" -ForegroundColor Green
Write-Host "  |  MinervaDB XL View Installation Complete!   |" -ForegroundColor Green
Write-Host "  +----------------------------------------------+" -ForegroundColor Green
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor White
Write-Host "  1. Edit config\config.ini with your PostgreSQL credentials"
Write-Host "  2. Activate the virtual environment:"
Write-Host "       .venv\Scripts\Activate.ps1"
Write-Host "  3. Test the connection:"
Write-Host "       python python\pg_connector.py"
Write-Host "  4. Import VBA modules into Excel (Alt+F11 -> File -> Import)"
Write-Host "  5. Run ETL pipeline:"
Write-Host "       python python\etl_pipeline.py --pipeline all"
Write-Host ""
Write-Host "  Full guide: docs\WINDOWS_INSTALL.md" -ForegroundColor Cyan
Write-Host "  Troubleshooting: docs\TROUBLESHOOTING.md" -ForegroundColor Cyan
Write-Host ""
