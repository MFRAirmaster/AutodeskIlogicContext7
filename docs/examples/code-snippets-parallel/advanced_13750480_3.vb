' Title: Apprentice Server with powershell, working prefix copy design
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apprentice-server-with-powershell-working-prefix-copy-design/td-p/13750480#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:28.496377

AddReference "System.IO"
AddReference "System.Windows.Forms"
AddReference "System.Drawing"
Imports Inventor
Imports IO = System.IO
Imports WF = System.Windows.Forms
Imports SD = System.Drawing
Imports System.Text
Imports System.Diagnostics

Public Class CopyDesignRule

  Public Sub Main()
    ' 1) Source assembly path from the active document
    Dim asmPath As String = ThisDoc.ModelDocument.FullFileName
    If String.IsNullOrEmpty(asmPath) OrElse Not IO.File.Exists(asmPath) Then
      WF.MessageBox.Show("Active document path not found or file does not exist.", _
                         "Copy Design", WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Error)
      Exit Sub
    End If

    ' 2) Build the form
    Dim frm As New WF.Form With {
      .Text            = "Copy Design (PowerShell)",
      .AutoScaleMode   = WF.AutoScaleMode.Dpi,
      .Size            = New SD.Size(760, 340),
      .Font            = New SD.Font("Segoe UI", 10.0!),
      .StartPosition   = WF.FormStartPosition.CenterScreen,
      .FormBorderStyle = WF.FormBorderStyle.FixedDialog,
      .MaximizeBox     = False,
      .MinimizeBox     = False
    }
    AddHandler frm.KeyDown, Sub(s, E) If E.KeyCode = WF.Keys.Escape Then frm.Close()

    Dim y As Integer = 18
    Dim gap As Integer = 40

    ' — Source (readonly)
    frm.Controls.Add(New WF.Label With {.Text = "Source assembly:", .Location = New SD.Point(15, y), .AutoSize = True})
    Dim tbSrc As New WF.TextBox With {.Location = New SD.Point(180, y - 3), .Width = 530, .Text = asmPath, .ReadOnly = True}
    frm.Controls.Add(tbSrc)
    y += gap

    ' — New sub-folder name
    frm.Controls.Add(New WF.Label With {.Text = "New folder name:", .Location = New SD.Point(15, y), .AutoSize = True})
    Dim tbFolder As New WF.TextBox With {.Location = New SD.Point(180, y - 3), .Width = 300, .Text = "CopyDesign"}
    frm.Controls.Add(tbFolder)
    y += gap

    ' — Destination root (optional)
    frm.Controls.Add(New WF.Label With {.Text = "Destination root (optional):", .Location = New SD.Point(15, y), .AutoSize = True})
    Dim tbRoot As New WF.TextBox With {.Location = New SD.Point(180, y - 3), .Width = 530, .Text = ""}
    frm.Controls.Add(tbRoot)
    y += gap

    ' — Copy prefix
    frm.Controls.Add(New WF.Label With {.Text = "Copy prefix (e.g. Copy_):", .Location = New SD.Point(15, y), .AutoSize = True})
    Dim tbPrefix As New WF.TextBox With {.Location = New SD.Point(180, y - 3), .Width = 180, .Text = "Copy_"}
    frm.Controls.Add(tbPrefix)
    y += gap

    ' — Existing prefixes to allow (comma-separated)
    frm.Controls.Add(New WF.Label With {
        .Text = "Existing prefixes to allow (comma-separated, e.g. WIP_,WIP2_):",
        .Location = New SD.Point(15, y), .AutoSize = True})
    Dim tbExtra As New WF.TextBox With {.Location = New SD.Point(15, y + 22), .Width = 695, .Text = ""}
    frm.Controls.Add(tbExtra)
    y += 70

    ' — Buttons
    Dim btnStart As New WF.Button With {.Text = "Start", .Location = New SD.Point(520, y), .Width = 90, .DialogResult = WF.DialogResult.OK}
    Dim btnClose As New WF.Button With {.Text = "Close", .Location = New SD.Point(620, y), .Width = 90, .DialogResult = WF.DialogResult.Cancel}
    frm.Controls.AddRange({btnStart, btnClose})
    frm.AcceptButton = btnStart
    frm.CancelButton = btnClose

    If frm.ShowDialog() <> WF.DialogResult.OK Then Exit Sub

    ' 3) Validate inputs and compute destination folder
    Dim baseFolder As String = IO.Path.GetDirectoryName(asmPath)
    Dim newFolder  As String = tbFolder.Text.Trim()
    If newFolder = "" Then
      WF.MessageBox.Show("You must enter a folder name.", "Copy Design", WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Warning)
      Exit Sub
    End If

    Dim rootPath As String = tbRoot.Text.Trim()
    Dim destDir As String
    If rootPath = "" Then
      ' default: new sub-folder next to the source assembly
      destDir = IO.Path.Combine(baseFolder, newFolder)
    Else
      ' user-supplied root path
      destDir = IO.Path.Combine(rootPath, newFolder)
    End If

    Dim prefix As String = tbPrefix.Text.Trim()
    If prefix = "" Then prefix = "Copy_"

    Dim extraText As String = tbExtra.Text.Trim()
    Dim extraList As New List(Of String)
    If extraText <> "" Then
      For Each tok In extraText.Split(","c)
        Dim p = tok.Trim()
        If p <> "" Then extraList.Add(p)
      Next
    End If

    ' 4) Generate and run the latest PowerShell script
    GenerateAndRunPS(asmPath, destDir, prefix, extraList)
  End Sub



  Private Sub GenerateAndRunPS(asmPath As String, destDir As String, prefix As String, extraPrefixes As List(Of String))
    ' Build PowerShell array literal for extra prefixes, with safe quoting
    Dim psExtra As String
    If extraPrefixes IsNot Nothing AndAlso extraPrefixes.Count > 0 Then
      Dim esc = extraPrefixes.Select(Function(s) s.Replace("'", "''"))
      psExtra = "@('" & String.Join("','", esc) & "')"
    Else
      psExtra = "@()"
    End If

    ' Create temp .ps1
    Dim psFile As String = IO.Path.Combine(IO.Path.GetTempPath(), "CopyDesign_" & Guid.NewGuid().ToString("N") & ".ps1")
    Dim sb As New StringBuilder()

    ' NOTE: This is the latest baseline logic (no Read-Host; uses values from the form)
    With sb
      .AppendLine("$ErrorActionPreference = 'Stop'")
      .AppendLine("")
      .AppendLine("# Inputs from iLogic")
      .AppendLine("$asmPath  = '" & asmPath.Replace("'", "''") & "'")
      .AppendLine("$destDir  = '" & destDir.Replace("'", "''") & "'")
      .AppendLine("$prefix   = '" & prefix.Replace("'", "''") & "'")
      .AppendLine("$extraPrefixes = " & psExtra)
      .AppendLine("")
      .AppendLine("# ensure output folder exists")
      .AppendLine("if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }")
      .AppendLine("")
      .AppendLine("# ApprenticeServer for CAD-only copy")
      .AppendLine("$appr = New-Object -ComObject Inventor.ApprenticeServer")
      .AppendLine("$asm  = $appr.Open($asmPath)")
      .AppendLine("$save = $appr.FileSaveAs")
      .AppendLine("")
      .AppendLine("function Pref([string]$path) { $prefix + ([IO.Path]::GetFileName($path)) }")
      .AppendLine("")
      .AppendLine("# 1) queue the main assembly")
      .AppendLine("$save.AddFileToSave($asm, (Join-Path $destDir (Pref $asmPath)))")
      .AppendLine("")
      .AppendLine("# 2) queue parts & sub-assemblies, filtered")
      .AppendLine("foreach ($doc in $asm.AllReferencedDocuments) {")
      .AppendLine("  if ($doc.NeedsMigrating) { continue }")
      .AppendLine("  $fileName = [IO.Path]::GetFileName($doc.FullFileName)")
      .AppendLine("  $baseName = [IO.Path]::GetFileNameWithoutExtension($fileName)")
      .AppendLine("  $ext      = [IO.Path]::GetExtension($fileName).ToLowerInvariant()")
      .AppendLine("  $matchesPattern = $baseName -match '^0.*-\d+$'")
      .AppendLine("  $matchesExtra   = $false")
      .AppendLine("  foreach ($p in $extraPrefixes) { if ($p -and $baseName.StartsWith($p)) { $matchesExtra = $true; break } }")
      .AppendLine("  if (($ext -in '.iam','.ipt') -and ($matchesPattern -or $matchesExtra)) {")
      .AppendLine("    $out    = Join-Path $destDir (Pref $doc.FullFileName)")
      .AppendLine("    $outDir = Split-Path $out")
      .AppendLine("    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }")
      .AppendLine("    $save.AddFileToSave($doc, $out)")
      .AppendLine("  } else {")
      .AppendLine("    Write-Host ('Skipping ' + $fileName + ' (naming filter)') -ForegroundColor DarkYellow")
      .AppendLine("  }")
      .AppendLine("}")
      .AppendLine("")
      .AppendLine("# 3) perform CAD copy")
      .AppendLine("$save.ExecuteSaveCopyAs()")
      .AppendLine("")
      .AppendLine("# 4) copy drawings from disk, filtered")
      .AppendLine("$exts = '.idw','.dwg','.ipn'")
      .AppendLine("function CopyDrawings($folder) {")
      .AppendLine("  foreach ($e in $exts) {")
      .AppendLine("    Get-ChildItem -Path $folder -Filter ""*$e"" -File -ErrorAction SilentlyContinue |")
      .AppendLine("    Where-Object {")
      .AppendLine("      if ($_.Extension -ieq '.idw') {")
      .AppendLine("        $b = $_.BaseName")
      .AppendLine("        $okPattern = $b -match '^0.*\d$'")
      .AppendLine("        $okExtra   = $false")
      .AppendLine("        foreach ($p in $extraPrefixes) { if ($p -and $b.StartsWith($p)) { $okExtra = $true; break } }")
      .AppendLine("        $okPattern -or $okExtra")
      .AppendLine("      } else { $true }")
      .AppendLine("    } |")
      .AppendLine("    ForEach-Object { Copy-Item $_.FullName (Join-Path $destDir ($prefix + $_.Name)) -Force }")
      .AppendLine("  }")
      .AppendLine("}")
      .AppendLine("CopyDrawings (Split-Path $asmPath -Parent)")
      .AppendLine("foreach ($doc in @($asm) + $asm.AllReferencedDocuments) { CopyDrawings (Split-Path $doc.FullFileName -Parent) }")
      .AppendLine("")
      .AppendLine("# 4.5) remove Read-Only attribute")
      .AppendLine("Get-ChildItem -Path $destDir -Recurse -File | Where-Object { $_.Extension -in '.iam','.ipt','.idw' } | ForEach-Object {")
      .AppendLine("  if ($_.IsReadOnly) { $_.IsReadOnly = $false; Write-Host ('Removed ReadOnly: ' + $_.FullName) }")
      .AppendLine("}")
      .AppendLine("")
      .AppendLine("# 5) first relink pass via Inventor.Application (.Path)")
      .AppendLine("function RelinkDrawings {")
      .AppendLine("  param($folder, $filePrefix)")
      .AppendLine("  $invApp = [Runtime.InteropServices.Marshal]::GetActiveObject('Inventor.Application')")
      .AppendLine("  Get-ChildItem -Path $folder -Include '*.idw','*.dwg' -File | ForEach-Object {")
      .AppendLine("    $draw = $invApp.Documents.Open($_.FullName, $true)")
      .AppendLine("    foreach ($desc in $draw.ReferencedDocumentDescriptors) {")
      .AppendLine("      $old  = $desc.ReferencedFileDescriptor.FullFileName")
      .AppendLine("      $base = [IO.Path]::GetFileName($old)")
      .AppendLine("      $new  = Join-Path $folder ($filePrefix + $base)")
      .AppendLine("      if (Test-Path $new) { $desc.ReferencedFileDescriptor.Path = $new }")
      .AppendLine("    }")
      .AppendLine("    $draw.Update(); $draw.Save(); $draw.Close()")
      .AppendLine("  }")
      .AppendLine("}")
      .AppendLine("RelinkDrawings -folder $destDir -filePrefix $prefix")
      .AppendLine("")
      .AppendLine("# 6) open result folder")
      .AppendLine("Start-Process explorer.exe -ArgumentList $destDir")
      .AppendLine("")
      .AppendLine("# 7) full-session ReplaceReference relink with iLogic suppression")
      .AppendLine("Write-Host 'Preparing full-session relink with iLogic/events off…' -ForegroundColor Cyan")
      .AppendLine("$inv = [Runtime.InteropServices.Marshal]::GetActiveObject('Inventor.Application')")
      .AppendLine("try { $inv.DisplayAlerts = $true } catch {}")
      .AppendLine("try { $inv.SilentOperation = $true } catch {}")
      .AppendLine("")
      .AppendLine("try {")
      .AppendLine("  $iLogicVB = $inv.ApplicationAddIns.ItemById('{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}')")
      .AppendLine("  if (-not $iLogicVB.Activated) { $iLogicVB.Activate() }")
      .AppendLine("  $iLogicVB.Automation.RulesEnabled           = $false")
      .AppendLine("  $iLogicVB.Automation.RulesOnEventsEnabled   = $false")
      .AppendLine("  $iLogicVB.Automation.SilentOperation        = $true")
      .AppendLine("  $iLogicVB.Automation.CallingFromOutside     = $true")
      .AppendLine("  Write-Host 'iLogic VB automation mode ON.'")
      .AppendLine("} catch { Write-Warning ('VB iLogic disable failed: ' + $_.Exception.Message) }")
      .AppendLine("")
      .AppendLine("# build source->dest map")
      .AppendLine("$map = @{}")
      .AppendLine("$allDocs = New-Object System.Collections.Generic.List[Object]")
      .AppendLine("$queue   = New-Object System.Collections.Generic.Queue[Object]")
      .AppendLine("$queue.Enqueue($asm)")
      .AppendLine("while ($queue.Count -gt 0) {")
      .AppendLine("  $d = $queue.Dequeue()")
      .AppendLine("  if (-not ($allDocs -contains $d)) {")
      .AppendLine("    $allDocs.Add($d)")
      .AppendLine("    foreach ($child in $d.AllReferencedDocuments) { if ($child.FullFileName) { $queue.Enqueue($child) } }")
      .AppendLine("  }")
      .AppendLine("}")
      .AppendLine("foreach ($doc in $allDocs) {")
      .AppendLine("  if ($doc.FullFileName) { $map[$doc.FullFileName.ToLower()] = Join-Path $destDir ($prefix + [IO.Path]::GetFileName($doc.FullFileName)) }")
      .AppendLine("}")
      .AppendLine("foreach ($file in Get-ChildItem -Path $destDir -Filter ""$prefix*.idw"" -File) {")
      .AppendLine("  $orig = $file.Name.Substring($prefix.Length)")
      .AppendLine("  $map[(Join-Path (Split-Path $asmPath -Parent) $orig).ToLower()] = $file.FullName")
      .AppendLine("}")
      .AppendLine("")
      .AppendLine("Get-ChildItem -Path $destDir -Filter ""$prefix*.idw"" -Recurse -File | ForEach-Object {")
      .AppendLine("  Write-Host ('  • ' + $_.Name)")
      .AppendLine("  $dw = $inv.Documents.Open($_.FullName, $false)")
      .AppendLine("  foreach ($r in $dw.ReferencedFileDescriptors) {")
      .AppendLine("    $k = $r.FullFileName.ToLower()")
      .AppendLine("    if ($map.ContainsKey($k)) { $r.DocumentDescriptor.ReferencedFileDescriptor.ReplaceReference($map[$k]) }")
      .AppendLine("  }")
      .AppendLine("  $dw.Update(); $dw.Save2(); $dw.Close($true)")
      .AppendLine("}")
      .AppendLine("")
      .AppendLine("# 8) update Part Number iProperties")
      .AppendLine("Write-Host 'Updating Part Number iProperties…' -ForegroundColor Cyan")
      .AppendLine("$inv2 = [Runtime.InteropServices.Marshal]::GetActiveObject('Inventor.Application')")
      .AppendLine("Get-ChildItem -Path $destDir -Filter ""$prefix*"" -File | ForEach-Object {")
      .AppendLine("  Write-Host ('  • ' + $_.Name)")
      .AppendLine("  $doc2 = $inv2.Documents.Open($_.FullName, $false)")
      .AppendLine("  try {")
      .AppendLine("    $ps = $doc2.PropertySets.Item('Design Tracking Properties')")
      .AppendLine("    $ps.Item('Part Number').Value = [IO.Path]::GetFileNameWithoutExtension($_.Name)")
      .AppendLine("    $doc2.Update(); $doc2.Save2()")
      .AppendLine("  } catch {")
      .AppendLine("    Write-Warning ('  ! Part Number failed on ' + $_.Name + ': ' + $_.Exception.Message)")
      .AppendLine("  } finally { $doc2.Close($true) }")
      .AppendLine("}")
      .AppendLine("Write-Host 'Part Number update complete.' -ForegroundColor Green")
    End With

    IO.File.WriteAllText(psFile, sb.ToString())

    ' Launch PowerShell so progress/errors are visible
    Dim psi As New ProcessStartInfo() With {
      .FileName        = "powershell.exe",
      .Arguments       = "-NoProfile -ExecutionPolicy Bypass -File """ & psFile & """",
      .UseShellExecute = True,
      .WindowStyle     = ProcessWindowStyle.Normal
    }

    Dim proc As Process = Process.Start(psi)
    ' Optional: wait for completion
    ' proc.WaitForExit()

    ' (Optional) delete temp script afterward
    ' Try
    '   IO.File.Delete(psFile)
    ' Catch ex As Exception
    '   WF.MessageBox.Show("Could not delete temp script:" & vbCrLf & ex.Message, "Copy Design", WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Warning)
    ' End Try
  End Sub

End Class