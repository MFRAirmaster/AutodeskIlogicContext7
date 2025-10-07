' Title: Save and replace parts in an assembly using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/save-and-replace-parts-in-an-assembly-using-ilogic/td-p/9409910
' Category: advanced
' Scraped: 2025-10-07T13:33:31.407966

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MsgBox("This rule '" & iLogicVb.RuleName & "' only works for Assembly Documents.",vbOK, "WRONG DOCUMENT TYPE")
	Return
End If

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument

Dim oFileDlg As Inventor.FileDialog = Nothing
ThisApplication.CreateFileDialog(oFileDlg)
oFileDlg.Filter = "Autodesk Inventor Assembly Files (*.iam)|*.iam"
oFileDlg.InitialDirectory = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
oFileDlg.FileName = IO.Path.GetFileName(oADoc.FullFileName).TrimEnd("."c,"i"c,"a"c,"m"c)
oFileDlg.DialogTitle = "Specify New Name & Location For Copied Assembly"
oFileDlg.CancelError = True

On Error Resume Next
oFileDlg.ShowSave
If Err.Number <> 0 Then
	MsgBox("No File Saved.", vbOKOnly, "DIALOG CANCELED")
ElseIf oFileDlg.FileName <> "" Then
	oNewFileName = oFileDlg.FileName
	oADoc.SaveAs(oNewFileName, False)
End If

oADoc = Nothing

InventorVb.DocumentUpdate()

oADoc = ThisApplication.ActiveDocument

Dim oLast3Chars As String
For Each oRefDoc As Document In oADoc.AllReferencedDocuments
	ThisApplication.Documents.Open(oRefDoc.FullFileName,False)
	oLast3Chars = Left(Right(oRefDoc.FullFileName,7),3)
	oRefDoc.SaveAs(oADoc.FullFileName & oLast3Chars, True)
	oRefDoc.Close
Next

Dim oOccDoc As Document
Dim oOccNewFileName As String
For Each oOcc As ComponentOccurrence In oADoc.ComponentDefinition.Occurrences
	oOccDoc = oOcc.Definition.Document
	oLast3Chars = Left(Right(oOccDoc.FullFileName,7),3)
	oOccNewFileName = Left(oADoc.FullFileName,Len(oADoc.FullFileName)-4) & oLast3Chars
	oOcc.Replace(oOccNewFileName, True)
Next