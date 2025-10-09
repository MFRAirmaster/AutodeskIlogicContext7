' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:12.526503

Sub Main
	If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
		MsgBox("An Assembly Document must be active for this rule to work. Exiting.",vbOKOnly+vbCritical, "WRONG DOCUMENT TYPE")
		Exit Sub
	End If
	Dim oADoc As AssemblyDocument = ThisAssembly.Document
	'save the 'active' assembly as a STEP file, using our Sub below
	SaveAsSTEP(oADoc)
	'create a List to keep track of which ones we have already exported
	'so we don't export the same document more than once
	Dim oExportedDocs As List(Of String)
	'this will process all parts & sub-assemlies in all levels of the main assembly
	StepDown(oADoc.ComponentDefinition.Occurrences, oExportedDocs)
	MsgBox("Finished saving files to STEP format.", vbInformation, "")
End Sub

Sub StepDown(oComps As ComponentOccurrences, oList As List(Of String))
	'this will process all parts & sub-assemlies in (top level only) of main assembly
	For Each oComp As ComponentOccurrence In oComps
		Dim oCompDoc As Document = oComp.Definition.Document
		If Not oList.Contains(oCompDoc.FullFileName) Then
			'save this component document as a STEP file, using our Sub below
			SaveAsSTEP(oCompDoc)
			'add the full file name of that document to the list
			oList.Add(oCompDoc.FullFileName)
		End If
		If TypeOf oComp.Definition Is AssemblyComponentDefinition Then
			StepDown(oComp.Definition.Occurrences, oList)
		End If
	Next
End Sub

Sub SaveAsSTEP(oDoc As Document)
	oDirChar = System.IO.Path.DirectorySeparatorChar
	oPath = System.IO.Path.GetDirectoryName(oDoc.FullFileName) & oDirChar & "STEP Files" & oDirChar
	If Not System.IO.Directory.Exists(oPath) Then System.IO.Directory.CreateDirectory(oPath)
	Dim oPN As String = oDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
	oNewName = oPath & oPN & ".stp"
	'Check to see if the STEP file already exists, if it does, ask if you want to overwrite it or not.
	If System.IO.File.Exists(oNewName) Then
		oAns = MsgBox("A STEP file with this name already exists." & vbCrLf & _
		"Do you want to overwrite it with this new one?",vbYesNo + vbQuestion + vbDefaultButton2, "FILE ALREADY EXISTS")
		If oAns = vbNo Then Exit Sub
	End If
	Try
		oDoc.SaveAs(oNewName, True)
	Catch
		MsgBox("Failed to save '" & oDoc.FullFileName & " out as a STEP file.", , "")
	End Try
End Sub