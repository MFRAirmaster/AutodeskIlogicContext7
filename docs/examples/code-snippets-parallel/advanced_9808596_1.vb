' Title: ilogic rule that creating a cope of a drawing and changing its iPart instance.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-that-creating-a-cope-of-a-drawing-and-changing-its/td-p/9808596#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:09:37.418034

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MsgBox("This rule '" & iLogicVb.RuleName & "' only works for Assembly Documents.",vbOKOnly, "WRONG DOCUMENT TYPE")
	Exit Sub 'or Return
End If

Dim oADoc As AssemblyDocument = ThisAssembly.Document
Dim oRefDoc As Document
'get an iPart within the assembly
For Each oRefDoc In oADoc.AllReferencedDocuments
	If oRefDoc.DocumentType = DocumentTypeEnum.kDesignElementDocumentObject Then Exit For
Next
'get the path where this iPart's file is found.
Dim oPath As String = IO.Path.GetDirectoryName(oRefDoc.FullFileName) & "\"

'define which other iPart you want to be represented within the new drawing.
Dim oNewModelName As String = "NewModel"
Dim oNewPart As String = oPath & oNewModelName & ".ipt"
Dim oNewDrawing As String = oPath & oNewModelName & ".idw"

If IO.File.Exists(oRefDoc.FullFileName.Replace(".ipt", ".idw")) Then
	'open the existing drawing for that iPart (assuming it has the same name and stored in the same directory)
	Dim oDDoc As DrawingDocument = ThisApplication.Documents.Open(oRefDoc.FullFileName.Replace(".ipt", ".idw"), False)
	'save a new copy of the file to disk, then close this old file
	oDDoc.SaveAs(oNewDrawing, True) 'saves a new copy of it to disk, while leaving the old file open
	oDDoc.Close(True) 'True  = skip save (this is the old document, so we don't want to save it.
	'now open the new copy of the drawing
	Dim oNewDDoc As DrawingDocument = ThisApplication.Documents.Open(oNewDrawing, False)
	'get the current file descriptor being referenced by this drawing
	Dim oFD As FileDescriptor = oNewDDoc.ReferencedDocumentDescriptors.Item(1).ReferencedFileDescriptor
	'replace it with the full file name of our new iPart's file.
	oFD.ReplaceReference(oNewPart)
	'update the drawing document
	oNewDDoc.Update()
	oNewDDoc.Save
	oNewDDoc.Close(True)
End If