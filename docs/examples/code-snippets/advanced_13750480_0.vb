' Title: Apprentice Server with powershell, working prefix copy design
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apprentice-server-with-powershell-working-prefix-copy-design/td-p/13750480#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:15:24.493088

# --------------------------------------------------------------------------------
# CopyDesign Script
#
# DESCRIPTION
#   Copies an Inventor assembly (.iam), its referenced parts (.ipt/.iam), and
#   drawing files (.idw/.dwg/.ipn) into a new folder, renames them with a user-
#   specified prefix, relinks all references, suppresses iLogic rule execution,
#   and updates the Part Number iProperty on each copied file.
#
# NAMING CONVENTION REQUIRED
#   • Model files (.iam, .ipt):
#       – Base name must start with '0' and end with '-<digits>' 
#         e.g. 0001-1.ipt, 0001-501.iam
#       – OR base name must start with one of your custom prefixes 
#         (e.g. WIP_, WIP2_)
#   • Drawing files (.idw):
#       – Base name must start with '0' and end with a digit 
#         e.g. 0001.idw
#       – OR start with one of your custom prefixes
#   • Other drawing types (.dwg, .ipn) are always copied regardless of name.
#   • Enter any extra prefixes when prompted; leave blank to skip.
#
# NOTE
#   Inventor must be running. The script attaches to the existing session.
# --------------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

#
# ←—— USER INPUT Prompts ————————————————————————————————————————
#

# Prompt for the folder containing the top-level assembly
$asmDir = Read-Host 'Enter the directory containing the top-level assembly'
if (-not (Test-Path $asmDir)) {
    Write-Error "Directory not found: $asmDir"
    exit 1
}

# Prompt for the assembly filename (with or without .iam)
$inputName = Read-Host 'Enter the top-level assembly filename (with or without .iam extension)'
if ($inputName -match '\.iam$') {
    $asmName = $inputName
} else {
    $asmName = "$inputName.iam"
}

# Build the full assembly path
$asmPath = Join-Path $asmDir $asmName
if (-not (Test-Path $asmPath)) {
    Write-Error "Assembly not found: $asmPath"
    exit 1
}

# Prompt for the destination folder
$destDir = Read-Host 'Enter the destination directory for copied files'
if (-not (Test-Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir | Out-Null
}

# Prompt for the filename prefix
$prefix = Read-Host 'Enter the prefix to prepend to copied filenames (e.g. Copy_)'
if (-not $prefix) {
    Write-Error "Prefix cannot be empty."
    exit 1
}

#
# ←—— Existing logic below unchanged, using $asmPath, $destDir, $prefix ——————————————————
#

# Prompt for any existing filename prefixes (e.g. WIP_, WIP2_) that should be allowed
# even if they don’t follow the “0…-digit” naming rule
$custom = Read-Host 'Enter any any existing filename prefixes if they do not follow the naming rule, to include in the copy, (comma-separated for different ones), or press Enter to skip'
if ($custom) { $extraPrefixes = $custom -split ',' | ForEach-Object { $_.Trim() } } else { $extraPrefixes = @() }

# ApprenticeServer for CAD-only copy
$appr = New-Object -ComObject Inventor.ApprenticeServer
$asm  = $appr.Open($asmPath)
$save = $appr.FileSaveAs

function Pref([string]$path) {
    $prefix + ([IO.Path]::GetFileName($path))
}

# 1) queue the main assembly
$save.AddFileToSave(
    $asm,
    (Join-Path $destDir (Pref $asmPath))
)

# 2) queue filtered parts & sub-assemblies
foreach ($doc in $asm.AllReferencedDocuments) {
    if ($doc.NeedsMigrating) { continue }

    $fileName = [IO.Path]::GetFileName($doc.FullFileName)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fileName)
    $ext      = [IO.Path]::GetExtension($fileName).ToLowerInvariant()

    $matchesPattern = $baseName -match '^0.*-\d+$'
    $matchesExtra   = $false
    foreach ($p in $extraPrefixes) {
        if ($p -and $baseName.StartsWith($p)) { $matchesExtra = $true; break }
    }

    if (($ext -in '.iam','.ipt') -and ($matchesPattern -or $matchesExtra)) {
        $out    = Join-Path $destDir (Pref $doc.FullFileName)
        $outDir = Split-Path $out
        if (-not (Test-Path $outDir)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        }
        $save.AddFileToSave(
            $doc,
            $out
        )
    }
    else {
        Write-Host "Skipping $fileName (naming filter)" -ForegroundColor DarkYellow
    }
}

# 3) perform CAD copy
$save.ExecuteSaveCopyAs()

# 4) copy drawings from disk, filtered
$exts = '.idw','.dwg','.ipn'
function CopyDrawings($folder) {
    foreach ($e in $exts) {
        Get-ChildItem -Path $folder -Filter "*$e" -File -ErrorAction SilentlyContinue |
        Where-Object {
            if ($_.Extension -ieq '.idw') {
                $b = $_.BaseName
                $okPattern = $b -match '^0.*\d$'
                $okExtra   = $false
                foreach ($p in $extraPrefixes) {
                    if ($p -and $b.StartsWith($p)) { $okExtra = $true; break }
                }
                $okPattern -or $okExtra
            }
            else { $true }
        } |
        ForEach-Object {
            Copy-Item $_.FullName (Join-Path $destDir ($prefix + $_.Name)) -Force
        }
    }
}
CopyDrawings (Split-Path $asmPath -Parent)
foreach ($doc in @($asm) + $asm.AllReferencedDocuments) {
    CopyDrawings (Split-Path $doc.FullFileName -Parent)
}

# 4.5) remove Read-Only attribute
Get-ChildItem -Path $destDir -Recurse -File |
    Where-Object { $_.Extension -in '.iam','.ipt','.idw' } |
    ForEach-Object {
        if ($_.IsReadOnly) {
            $_.IsReadOnly = $false
            Write-Host "Removed ReadOnly: $($_.FullName)"
        }
    }

# 5) first relink pass via ApprenticeServer
function RelinkDrawings {
    param($folder, $filePrefix)
    $invApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Inventor.Application")
    Get-ChildItem -Path $folder -Include '*.idw','*.dwg' -File | ForEach-Object {
        $draw = $invApp.Documents.Open($_.FullName, $true)
        foreach ($desc in $draw.ReferencedDocumentDescriptors) {
            $old  = $desc.ReferencedFileDescriptor.FullFileName
            $base = [IO.Path]::GetFileName($old)
            $new  = Join-Path $folder ($filePrefix + $base)
            if (Test-Path $new) {
                $desc.ReferencedFileDescriptor.Path = $new
            }
        }
        $draw.Update(); $draw.Save(); $draw.Close()
    }
}
RelinkDrawings -folder $destDir -filePrefix $prefix

# 6) open result folder
Start-Process explorer.exe -ArgumentList $destDir

# 7) full-session ReplaceReference relink (suppress pop-ups & iLogic)
Write-Host "Preparing full-session relink with iLogic/events off…" -ForegroundColor Cyan

# attach to running Inventor
$inv = [Runtime.InteropServices.Marshal]::GetActiveObject("Inventor.Application")

# allow alerts, but run silently
try { $inv.DisplayAlerts  = $true  } catch {}
try { $inv.SilentOperation = $true  } catch {}

# disable VB iLogic automation mode
try {
    $iLogicVB = $inv.ApplicationAddIns.ItemById('{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}')
    if (-not $iLogicVB.Activated) { $iLogicVB.Activate() }
    $iLogicVB.Automation.RulesEnabled           = $false
    $iLogicVB.Automation.RulesOnEventsEnabled   = $false
    $iLogicVB.Automation.SilentOperation        = $true
    $iLogicVB.Automation.CallingFromOutside     = $true
    Write-Host "iLogic VB automation mode ON."
}
catch { Write-Warning "VB iLogic disable failed: $($_.Exception.Message)" }

# build source->dest map
$map = @{}
$allDocs = New-Object System.Collections.Generic.List[Object]
$queue   = New-Object System.Collections.Generic.Queue[Object]
$queue.Enqueue($asm)
while ($queue.Count -gt 0) {
    $d = $queue.Dequeue()
    if (-not ($allDocs -contains $d)) {
        $allDocs.Add($d)
        foreach ($child in $d.AllReferencedDocuments) {
            if ($child.FullFileName) { $queue.Enqueue($child) }
        }
    }
}
foreach ($doc in $allDocs) {
    if ($doc.FullFileName) {
        $map[$doc.FullFileName.ToLower()] = Join-Path $destDir ($prefix + [IO.Path]::GetFileName($doc.FullFileName))
    }
}
foreach ($file in Get-ChildItem -Path $destDir -Filter "$prefix*.idw" -File) {
    $orig = $file.Name.Substring($prefix.Length)
    $map[(Join-Path (Split-Path $asmPath -Parent) $orig).ToLower()] = $file.FullName
}

# ReplaceReference relink loop
Get-ChildItem -Path $destDir -Filter "$prefix*.idw" -Recurse -File | ForEach-Object {
    Write-Host "  • $($_.Name)"
    $dw = $inv.Documents.Open($_.FullName, $false)
    foreach ($r in $dw.ReferencedFileDescriptors) {
        $k = $r.FullFileName.ToLower()
        if ($map.ContainsKey($k)) {
            $r.DocumentDescriptor.ReferencedFileDescriptor.ReplaceReference($map[$k])
        }
    }
    $dw.Update(); $dw.Save2(); $dw.Close($true)
}

# 8) update Part Number iProperties
Write-Host "Updating Part Number iProperties…" -ForegroundColor Cyan
$inv2 = [Runtime.InteropServices.Marshal]::GetActiveObject("Inventor.Application")
Get-ChildItem -Path $destDir -Filter "$prefix*" -File | ForEach-Object {
    Write-Host "  • $($_.Name)"
    $doc2 = $inv2.Documents.Open($_.FullName, $false)
    try {
        $ps = $doc2.PropertySets.Item("Design Tracking Properties")
        $ps.Item("Part Number").Value = [IO.Path]::GetFileNameWithoutExtension($_.Name)
        $doc2.Update(); $doc2.Save2()
    }
    catch {
        Write-Warning "  ! Part Number failed on $($_.Name): $($_.Exception.Message)"
    }
    finally {
        $doc2.Close($true)
    }
}
Write-Host "Part Number update complete." -ForegroundColor Green