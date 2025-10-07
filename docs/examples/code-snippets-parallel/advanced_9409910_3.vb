' Title: Save and replace parts in an assembly using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/save-and-replace-parts-in-an-assembly-using-ilogic/td-p/9409910#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:58:15.609350

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MsgBox("This rule '" & iLogicVb.RuleName & "' only works for Assembly Documents.",vbOK, "WRONG DOCUMENT TYPE")
	Return
End If

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument

Dim oFileDlg As FileDialog = Nothing
ThisApplication.CreateFileDialog(oFileDlg)
oFileDlg.Filter = "Autodesk Inventor Assembly Files (*.iam)|*.iam"
oFileDlg.InitialDirectory = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
oFileDlg.FileName = IO.Path.GetFileName(oADoc.FullFileName)
oFileDlg.Title = "Specify New Name & Location For Copied Assembly"
On Error Resume Next
oFileDlg.ShowDialog
If Err.Number <> 0 Then
	MsgBox("No File Saved.", vbOKOnly, "DIALOG CANCELED")
ElseIf oFileDlg.FileName <> "" Then
	oNewFileName = oFileDlg.FileName
	oADoc.SaveAs(oNewFileName, False)
End If

oADoc = Nothing

InventorVb.DocumentUpdate()

oADoc = ThisApplication.ActiveDocument

For Each oRefDoc As Document In oADoc.AllReferencedDocuments
	Dim oEndDigits As String = Right(oRefDoc.FullFileName,2)
	oRefDoc.SaveAs(oADoc.FullFileName & oEndDigits, True)
Next

For Each oOcc As ComponentOccurrence In oADoc.ComponentDefinition.Occurrences
	Dim oOccDoc As Document = oOcc.Definition.Document
	Dim oOccFileName As String = oOccDoc.FullFileName
	Dim oOccPath As String = IO.Path.GetFullPath(oOccFileName)
	Dim oOccNewFileName As String = oOccDoc.FullFileName
	'oOcc.Replace(oOcc.N, True)
Next