#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Autodesk Forum Scraper for iLogic Code Examples
    
.DESCRIPTION
    PowerShell wrapper for the advanced forum scraper that allows you to scrape
    the Autodesk Inventor Programming Forum for iLogic code examples.
    
.PARAMETER MaxPages
    Maximum number of forum pages to scrape (default: 10)
    Each page typically contains 20-30 posts
    
.PARAMETER Delay
    Delay in seconds between requests (default: 2.0)
    Increase this if you're getting rate-limited
    
.PARAMETER MaxPosts
    Maximum number of individual posts to scrape (optional)
    If not specified, will scrape all posts found in MaxPages
    
.PARAMETER OutputDir
    Directory for output files (default: docs/examples/advanced-forum-scraped)
    
.PARAMETER InstallDeps
    Install required Python dependencies before running
    
.EXAMPLE
    .\Run-ForumScraper.ps1 -MaxPages 5
    Scrape 5 forum pages
    
.EXAMPLE
    .\Run-ForumScraper.ps1 -MaxPages 20 -Delay 3
    Scrape 20 pages with 3 second delay between requests
    
.EXAMPLE
    .\Run-ForumScraper.ps1 -MaxPages 10 -MaxPosts 50
    Scrape up to 10 pages but stop after 50 posts with code
    
.EXAMPLE
    .\Run-ForumScraper.ps1 -InstallDeps
    Install dependencies first, then scrape with default settings
#>

param(
    [Parameter(Mandatory=$false)]
    [int]$MaxPages = 10,
    
    [Parameter(Mandatory=$false)]
    [double]$Delay = 2.0,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxPosts = 0,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "docs/examples/advanced-forum-scraped",
    
    [Parameter(Mandatory=$false)]
    [switch]$InstallDeps
)

# Script configuration
$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
$PythonScript = Join-Path $ScriptDir "advanced-forum-scraper.py"

# Color output functions
function Write-Header {
    param([string]$Message)
    Write-Host "`n================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "================================`n" -ForegroundColor Cyan
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
        
        Write-Info "Installing Playwright browsers..."
        & $PythonCmd -m playwright install chromium
        
        Write-Success "Playwright browsers installed"
    } catch {
        Write-Error "Failed to install dependencies: $_"
        exit 1
    }
}

# Create temporary Python script with parameters
function New-ConfiguredScript {
    param(
        [int]$MaxPages,
        [double]$Delay,
        [int]$MaxPosts,
        [string]$OutputDir
    )
    
    Write-Info "Creating configured scraper script..."
    
    # Read the original script
    $scriptContent = Get-Content $PythonScript -Raw
    
    # Create a temporary modified version with custom parameters
    $tempScript = Join-Path $env:TEMP "temp_forum_scraper_$(Get-Date -Format 'yyyyMMddHHmmss').py"
    
    # Modify the main() function to use our parameters
    $modifiedContent = $scriptContent -replace 'async def main\(\):.*?await scraper\.run\(max_pages=\d+, delay=[\d.]+\)', @"
async def main():
    ""${'"'}Main entry point""${'"'}
    scraper = AdvancedForumScraper()
    scraper.output_dir = r'$OutputDir'
    scraper.code_dir = r'$OutputDir\..\code-snippets'
    
    # Run with custom settings
    await scraper.run(max_pages=$MaxPages, delay=$Delay)
"@
    
    # If MaxPosts is specified, add post limit logic
    if ($MaxPosts -gt 0) {
        $modifiedContent = $modifiedContent -replace '# Scrape individual posts', @"
# Scrape individual posts (max: $MaxPosts)
            posts_with_code = 0
            max_posts_limit = $MaxPosts
"@
        
        $modifiedContent = $modifiedContent -replace 'if post_data:', @"
if post_data:
                    await self.save_post_data(post_data)
                    self.stats['with_code'] += 1
                    posts_with_code += 1
                    
                    if posts_with_code >= max_posts_limit:
                        print(f"`n🎯 Reached maximum posts limit ({max_posts_limit})")
                        break
                else:
                    if post_data:
"@
    }
    
    $modifiedContent | Set-Content $tempScript -Encoding UTF8
    
    return $tempScript
}

# Main execution
function Main {
    Write-Header "Autodesk Forum Scraper for iLogic"
    
    Write-Info "Configuration:"
    Write-Host "  Max Pages:     $MaxPages" -ForegroundColor White
    Write-Host "  Delay:         ${Delay}s" -ForegroundColor White
    if ($MaxPosts -gt 0) {
        Write-Host "  Max Posts:     $MaxPosts" -ForegroundColor White
    } else {
        Write-Host "  Max Posts:     Unlimited" -ForegroundColor White
    }
    Write-Host "  Output Dir:    $OutputDir" -ForegroundColor White
    
    # Check Python
    $pythonCmd = Test-Python
    
    # Install dependencies if requested
    if ($InstallDeps) {
        Install-Dependencies -PythonCmd $pythonCmd
    }
    
    # Verify script exists
    if (-not (Test-Path $PythonScript)) {
        Write-Error "Python scraper script not found: $PythonScript"
        exit 1
    }
    
    # Create output directory
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
        Write-Success "Created output directory: $OutputDir"
    }
    
    # Create configured script
    $tempScript = New-ConfiguredScript -MaxPages $MaxPages -Delay $Delay -MaxPosts $MaxPosts -OutputDir $OutputDir
    
    try {
        Write-Header "Starting Scraper"
        Write-Info "This may take several minutes depending on settings..."
        Write-Info "Press Ctrl+C to stop the scraper"
        
        # Run the scraper
        & $pythonCmd $tempScript
        
        if ($LASTEXITCODE -eq 0) {
            Write-Success "Scraping completed successfully!"
            Write-Info "Results saved to: $OutputDir"
            
            # Show summary
            Write-Header "Summary"
            
            # Count generated files
            $mdFiles = Get-ChildItem -Path $OutputDir -Filter "*.md" -ErrorAction SilentlyContinue
            $vbFiles = Get-ChildItem -Path (Join-Path $OutputDir "..\code-snippets") -Filter "*.vb" -ErrorAction SilentlyContinue
            
            if ($mdFiles) {
                Write-Success "Generated $($mdFiles.Count) documentation files"
            }
            if ($vbFiles) {
                Write-Success "Extracted $($vbFiles.Count) code snippets"
            }
            
            Write-Info "You can now use these examples as context for iLogic development!"
        } else {
            Write-Error "Scraping failed with exit code: $LASTEXITCODE"
            exit $LASTEXITCODE
        }
    } catch {
        Write-Error "An error occurred: $_"
        exit 1
    } finally {
        # Cleanup temporary script
        if (Test-Path $tempScript) {
            Remove-Item $tempScript -Force -ErrorAction SilentlyContinue
        }
    }
}

# Run the script
Main
