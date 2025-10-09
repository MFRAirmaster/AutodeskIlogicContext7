' Title: dxfAddIn.savecopyAs uses old dxfoptions used in manual saving
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/dxfaddin-savecopyas-uses-old-dxfoptions-used-in-manual-saving/td-p/13836969#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:09:03.304804

If dxfAddIn.HasSaveCopyAsOptions(oDoc, dxfContext, dxfOptions) Then
    Dim configPath As String = "C:\InventorStandaard\iLogic\Standard_dxf\DXFconfigurationfile.ini"

    ' ==== Verify config path exists ====
    If Not System.IO.File.Exists(configPath) Then
        MsgBox("DXF config file not found at: " & vbCrLf & configPath, vbCritical, "Export Cancelled")
        Exit Sub   ' stop the rule if .ini is missing
    End If

    ' Clear old cached options first (important!)
    dxfOptions.Clear()

    ' Force both keys that the translator looks for
    dxfOptions.Value("ConfigurationFile") = configPath
    dxfOptions.Value("Export_AcadIniFile") = configPath
    dxfOptions.Value("ModelGeometryOnly") = True

    ' Confirm ConfigurationFile was actually set
    Dim configVal As String = ""
    Try
        configVal = dxfOptions.Value("ConfigurationFile")
    Catch
        configVal = "(ConfigurationFile not set!)"
    End Try

    ' Build a text list of all DXF options for debug
    Dim report As String = "DXF Export Options:" & vbCrLf
    For i As Integer = 1 To dxfOptions.Count
        Dim key As String = dxfOptions.Name(i)
        Dim val As String = ""
        Try
            val = dxfOptions.Value(key)
        Catch
            val = "(no value)"
        End Try
        report &= "  " & key & " = " & val & vbCrLf
    Next

    MsgBox("DXF Config file now set to: " & configVal & vbCrLf & vbCrLf & report, _
           vbInformation, "DXF Options Debug")
End If

' Perform the export
Try
    dxfAddIn.SaveCopyAs(oDoc, dxfContext, dxfOptions, dxfData)
Catch ex As Exception
    MessageBox.Show("Error during DXF export: " & ex.Message, "Export Error")
End Try

' ==== Done ====
MsgBox("PDF and DXF exported to specified folders successfully." & vbCrLf & _
       "PDF: " & pdfPath & vbCrLf & "DXF: " & dxfPath, vbInformation, "Export Complete")