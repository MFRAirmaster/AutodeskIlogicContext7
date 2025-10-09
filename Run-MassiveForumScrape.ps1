#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Massive Autodesk Forum Scraper - 200+ Pages
    
.DESCRIPTION
    PowerShell wrapper to run the parallel forum scraper for massive data collection.
    Scrapes 200+ pages with ~57 posts per page (~11,400+ posts total).
    All code snippets are saved to database and individual files.
    
.PARAMETER MaxPages
    Maximum number of forum pages to scrape (default: 250)
    Each page typically contains ~57 posts
    
.PARAMETER MaxWorkers
    Number of parallel workers (default: 8)
    
.PARAMETER InstallDeps
    Install required Python dependencies before running
    
.EXAMPLE
    .\Run-MassiveForumScrape.ps1
    Scrape 250 pages with default settings
    
.EXAMPLE
    .\Run-MassiveForumScrape.ps1 -MaxPages 300 -MaxWorkers 10
    Scrape 300 pages with 10 parallel workers
    
.EXAMPLE
    .\Run-MassiveForumScrape.ps1 -InstallDeps
    Install dependencies first, then scrape with default settings
#>

param(
    [Parameter(Mandatory=$false)]
    [int]$MaxPages = 250,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxWorkers = 8,
    
    [Parameter(Mandatory=$false)]
    [switch]$InstallDeps
)

# Script configuration
$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
$PythonScript = Join-Path $ScriptDir "parallel-forum-scraper.py"

# Color output functions
function Write-Header {
    param([string]$Message)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "ℹ️  $Message" -ForegroundColor Blue
}

function Write-Warning {
    param([string]$Message)
    Write-Host "⚠️  $Message" -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Message)
    Write-Host "❌ $Message" -ForegroundColor Red
}

# Check Python installation
function Test-Python {
    Write-Info "Checking Python installation..."
    
    $pythonCommands = @("python", "python3", "py")
    $pythonCmd = $null
    
    foreach ($cmd in $pythonCommands) {
        try {
            $version = & $cmd --version 2>&1
            if ($LASTEXITCODE -eq 0) {
                $pythonCmd = $cmd
                Write-Success "Found Python: $version"
                break
            }
        } catch {
            continue
        }
    }
    
    if (-not $pythonCmd) {
        Write-Error "Python not found. Please install Python 3.8 or higher."
        Write-Info "Download from: https://www.python.org/downloads/"
        exit 1
    }
    
    return $pythonCmd
}

# Install Python dependencies
function Install-Dependencies {
    param([string]$PythonCmd)
    
    Write-Header "Installing Python Dependencies"
    
    $packages = @(
        "playwright",
        "aiofiles"
    )
    
    Write-Info "Installing required packages: $($packages -join ', ')"
    
    try {
        & $PythonCmd -m pip install --upgrade pip
        & $PythonCmd -m pip install $packages
        
        Write-Success "Python packages installed"
        
        Write-Info "Installing Playwright browsers (this may take a few minutes)..."
        & $PythonCmd -m playwright install chromium
        
        Write-Success "Playwright browsers installed"
    } catch {
        Write-Error "Failed to install dependencies: $_"
        exit 1
    }
}

# Main script
Write-Header "🚀 MASSIVE AUTODESK FORUM SCRAPER 🚀"

# Get Python command
$pythonCmd = Test-Python

# Install dependencies if requested
if ($InstallDeps) {
    Install-Dependencies -PythonCmd $pythonCmd
}

# Verify scraper script exists
if (-not (Test-Path $PythonScript)) {
    Write-Error "Scraper script not found: $PythonScript"
    exit 1
}

# Display scraping plan
Write-Header "Scraping Configuration"
Write-Info "Target Pages: $MaxPages pages"
Write-Info "Expected Posts: ~$($MaxPages * 57) posts (assuming ~57 posts per page)"
Write-Info "Parallel Workers: $MaxWorkers concurrent workers"
Write-Info "Output Directory: docs/examples/parallel-forum-scraped/"
Write-Info "Code Snippets: docs/examples/code-snippets-parallel/"
Write-Info "Database: docs/examples/ilogic-snippets.db"
Write-Host ""

# Estimate time
$estimatedMinutes = [math]::Ceiling(($MaxPages * 57) / ($MaxWorkers * 2))
Write-Warning "Estimated time: ~$estimatedMinutes minutes (depends on network and server)"
Write-Host ""

# Confirm before starting
$confirmation = Read-Host "Ready to start massive scraping? This will take a while. (Y/N)"
if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
    Write-Info "Scraping cancelled by user."
    exit 0
}

Write-Header "Starting Parallel Forum Scraping"
Write-Info "Script: $PythonScript"
Write-Info "This will scrape multiple forum sections in parallel..."
Write-Host ""

# Modify the Python script temporarily to use command-line args
$tempScript = Join-Path $env:TEMP "temp-parallel-scraper.py"
$scriptContent = Get-Content $PythonScript -Raw

# Replace the main() function to use our parameters
$scriptContent = $scriptContent -replace 'scraper = ParallelForumScraper\(max_workers=\d+\)', "scraper = ParallelForumScraper(max_workers=$MaxWorkers)"
$scriptContent = $scriptContent -replace 'await scraper\.run\(max_pages=\d+\)', "await scraper.run(max_pages=$MaxPages)"

# Save temp script
Set-Content -Path $tempScript -Value $scriptContent

try {
    # Run the scraper
    $startTime = Get-Date
    
    & $pythonCmd $tempScript
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Header "Scraping Complete!"
    Write-Success "Duration: $($duration.ToString('hh\:mm\:ss'))"
    Write-Info "Results saved to: docs/examples/parallel-forum-scraped/"
    Write-Info "Code snippets: docs/examples/code-snippets-parallel/"
    Write-Info "Database: docs/examples/ilogic-snippets.db"
    
    # Show database stats
    Write-Host ""
    Write-Info "Running database analysis..."
    
    $analyzeScript = Join-Path $ScriptDir "analyze_db.py"
    if (Test-Path $analyzeScript) {
        & $pythonCmd $analyzeScript
    }
    
} catch {
    Write-Error "Scraping failed: $_"
    exit 1
} finally {
    # Clean up temp file
    if (Test-Path $tempScript) {
        Remove-Item $tempScript -Force
    }
}

Write-Host ""
Write-Success "All done! Check the output directories for results."
Write-Info "Next step: Analyze the database with analyze_db.py for statistics"
