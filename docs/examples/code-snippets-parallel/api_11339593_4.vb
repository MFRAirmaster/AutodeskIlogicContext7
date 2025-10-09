' Title: Inventor DWG to AutoCAD DWG Conversion
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inventor-dwg-to-autocad-dwg-conversion/td-p/11339593#messageview_0
' Category: api
' Scraped: 2025-10-09T09:02:01.143165

Sub Main()
    If Not ConfirmProceed() Then Return

    Dim sourceFolder As String = GetFolder("Select folder containing DWG files")
    If sourceFolder = "" Then Return

    Dim exportFolder As String = GetFolder("Select folder to save AutoCAD DWG files")
    If exportFolder = "" Then Return

    Dim configPath As String = "C:\Users\thoma\Documents\$_INDUSTRIAL_DRAFTING_&_DESIGN\INVENTOR_FILES\PRESETS\EXPORT-DWG.ini"
    If Not System.IO.File.Exists(configPath) Then
        MessageBox.Show("Missing DWG config: " & configPath)
        Return
    End If

    Dim files() As String = System.IO.Directory.GetFiles(sourceFolder, "*.dwg")
    Dim successCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim errorLog As New System.Text.StringBuilder
    Dim exportLog As New System.Text.StringBuilder

    exportLog.AppendLine("DWG Export Log - " & Now.ToString())
    exportLog.AppendLine("")

    For Each file As String In files
        Dim doc As Document = Nothing
        Try
            doc = ThisApplication.Documents.Open(file, False)
            If doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                doc.Close(True)
                Continue For
            End If

            Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(file)
            Dim drawDoc As DrawingDocument = doc
            Dim sheet1 As Sheet = drawDoc.Sheets(1)
            If sheet1.DrawingViews.Count = 0 Then Throw New Exception("No views on Sheet 1")
            Dim baseView As DrawingView = sheet1.DrawingViews(1)

            Dim modelDoc As Document = baseView.ReferencedDocumentDescriptor.ReferencedDocument
            If modelDoc Is Nothing Then Throw New Exception("Referenced model is missing.")

            Dim modelRev As String = ""
            Try
                modelRev = iProperties.Value(modelDoc, "Project", "Revision Number")
            Catch
                modelRev = "NO-REV"
            End Try

            Dim exportName As String = baseName & "-REV-" & SanitizeFileName(modelRev) & ".dwg"
            Dim exportPath As String = System.IO.Path.Combine(exportFolder, exportName)

            exportLog.AppendLine("DWG: " & baseName)
            exportLog.AppendLine(" ↳ Model: " & modelDoc.DisplayName)
            exportLog.AppendLine(" ↳ REV: " & modelRev)

            If System.IO.File.Exists(exportPath) Then
                exportLog.AppendLine(" ⏭ Skipped (same REV already exported)")
                doc.Close(True)
                skippedCount += 1
                Continue For
            End If

            ' Set up translator
            Dim addIn As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
            Dim context = ThisApplication.TransientObjects.CreateTranslationContext
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim options = ThisApplication.TransientObjects.CreateNameValueMap
            If addIn.HasSaveCopyAsOptions(doc, context, options) Then
                options.Value("Export_Acad_IniFile") = configPath
            End If
            Dim oData = ThisApplication.TransientObjects.CreateDataMedium
            oData.FileName = exportPath

            addIn.SaveCopyAs(doc, context, options, oData)
            doc.Close(True)
            successCount += 1
            exportLog.AppendLine(" ✅ Exported: " & exportName)
        Catch ex As Exception
            If Not doc Is Nothing Then doc.Close(True)
            errorLog.AppendLine("ERROR processing " & file & ": " & ex.Message)
        End Try

        exportLog.AppendLine("")
    Next

    ' Write logs
    System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "DWG_Export_Log.txt"), exportLog.ToString())
    If errorLog.Length > 0 Then
        System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "DWG_Export_Errors.txt"), errorLog.ToString())
    End If

    MessageBox.Show("DWG export complete." & vbCrLf & _
        "✔ Files exported: " & successCount & vbCrLf & _
        "⏭ Skipped (already exported at this REV): " & skippedCount & vbCrLf & _
        "⚠ Errors: " & errorLog.Length, "Export Summary")
End Sub

Function GetFolder(prompt As String) As String
    Dim dlg As New FolderBrowserDialog
    dlg.Description = prompt
    If dlg.ShowDialog() <> DialogResult.OK Then Return ""
    Return dlg.SelectedPath
End Function

Function ConfirmProceed() As Boolean
    Dim result = MessageBox.Show("Export DWGs only if target REV file does not exist?", "Confirm Export", MessageBoxButtons.YesNo)
    Return result = DialogResult.Yes
End Function

Function SanitizeFileName(input As String) As String
    For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
        input = input.Replace(c, "_")
    Next
    Return input
End Function