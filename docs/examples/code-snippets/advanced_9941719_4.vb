' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719
' Category: advanced
' Scraped: 2025-10-07T13:17:40.528259

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MsgBox("An Assembly Document must be active for this rule to work. Exiting.",vbOKOnly+vbCritical, "WRONG DOCUMENT TYPE")
	Exit Sub
End If
Dim oADoc As AssemblyDocument = ThisAssembly.Document
Dim oAsmName As String = System.IO.Path.GetFileNameWithoutExtension(oADoc.FullFileName)
'Dim oAsmPN As String = oADoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
Dim oPath As String = System.IO.Path.GetDirectoryName(oADoc.FullFileName) & "\STEP Files\"
If Not System.IO.Directory.Exists(oPath) Then
	System.IO.Directory.CreateDirectory(oPath)
End If
oADoc.SaveAs(oPath & oAsmName & ".stp", True)
Dim oRefName, oRefPN As String
For Each oRefDoc As Document In oADoc.AllReferencedDocuments
	oRefDoc = ThisApplication.Documents.Open(oRefDoc.FullDocumentName, False)
	'oRefName = System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullFileName)
	oRefPN = oRefDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
	Try
		'oRefDoc.SaveAs(oPath & oRefName & ".stp", True)
		oRefDoc.SaveAs(oPath & oRefPN & ".stp", True)
	Catch
		MsgBox("Failed to save '" & oRefDoc.FullFileName & " out as an STEP file.", vbOKOnly, " ")
	End Try
	oRefDoc.Close(True) 'True = Skip Save
Next
MsgBox("All new STEP files were saved to:" & vbCrLf & oPath, vbOKOnly,"FINISHED")