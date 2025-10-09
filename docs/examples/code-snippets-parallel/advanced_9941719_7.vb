' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:12.526503

Sub Main
	If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
		MsgBox("An Assembly Document must be active for this rule to work. Exiting.",vbOKOnly+vbCritical, "WRONG DOCUMENT TYPE")
		Exit Sub
	End If
	Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
	oDirChar = System.IO.Path.DirectorySeparatorChar
	'define path to save STEP file to
	oPath = System.IO.Path.GetDirectoryName(oADoc.FullFileName) & oDirChar & "EXPORT_STEP" & oDirChar
	'make sure that path exists
	If Not System.IO.Directory.Exists(oPath) Then
		System.IO.Directory.CreateDirectory(oPath)
	End If
	'create a List to keep track of which ones we have already exported
	Dim oExportedDocs As List(Of String)
	'loop through the components in this assembly
	For Each oOcc As ComponentOccurrence In oADoc.ComponentDefinition.Occurrences
		'if it's not a part, skip to next component
		If oOcc.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For
		'get the document this component is representing
		Dim oOccPDoc As PartDocument = oOcc.Definition.Document
		'if this document is already in our list, we've already saved it out, so skip to next component
		If oExportedDocs.Contains(oOccPDoc.FullFileName) Then Continue For
		'get the file name (without path or extension) 
		oName = System.IO.Path.GetFileNameWithoutExtension(oOccPDoc.FullFileName)
		'combine path, name, & new extension
		oNewName = oPath & oName & ".stp"
		'Check to see if the STEP file already exists, if it does, ask if you want to overwrite it or not.
		If System.IO.File.Exists(oNewName) Then
			oAns = MsgBox("A STEP file with this name already exists." & vbCrLf & _
			"Do you want to overwrite it with this new one?", vbYesNo + vbQuestion + vbDefaultButton2, "FILE ALREADY EXISTS")
			If oAns = vbNo Then Continue For
		End If
		Try
			'save it out as STEP file
			oOccPDoc.SaveAs(oNewName, True)
			'add this document's full file name to our list
			oExportedDocs.Add(oOccPDoc.FullFileName)
		Catch oEx As Exception
			MsgBox("Failed to save '" & oOccPDoc.FullFileName & " out as a STEP file." & vbCrLf & _
			oEx.Message & vbCrLf & oEx.StackTrace, vbExclamation, "iLogic")
		End Try
	Next
	MsgBox("Finished saving files to STEP format.", vbInformation, "iLogic")
End Sub