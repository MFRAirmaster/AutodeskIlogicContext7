' Title: Save and Replace Rule
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/save-and-replace-rule/td-p/13802907#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:05:21.712644

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

' Get the new base name without extension and path
Dim oNewBaseName As String = IO.Path.GetFileNameWithoutExtension(oADoc.FullFileName)

' Extract the new prefix (everything before the first hyphen)
Dim oNewPrefix As String = ""
Dim oHyphenPos As Integer = oNewBaseName.IndexOf("-")
If oHyphenPos > 0 Then
    oNewPrefix = oNewBaseName.Substring(0, oHyphenPos)
End If

' Calculate how many characters to keep from original name (excluding "Project")
' Based on your example: "Project-100000" -> "250001-100000"
' We keep 8 characters total, minus 6 for "Project" = 6 characters after "Project"
Dim oCharsToKeepAfterProject As Integer = oNewBaseName.Length - oNewPrefix.Length

Dim oReferenceSuffix As String
For Each oRefDoc As Document In oADoc.AllReferencedDocuments
	ThisApplication.Documents.Open(oRefDoc.FullFileName,False)
	
	' Get original filename without extension
	Dim oOriginalName As String = IO.Path.GetFileNameWithoutExtension(oRefDoc.FullFileName)
	
	' Check if it starts with "Project" and extract the suffix
	If oOriginalName.StartsWith("Project") And oOriginalName.Length > 7 Then
		' Take the specified number of characters after "Project"
		Dim oAfterProject As String = oOriginalName.Substring(7)
		If oAfterProject.Length >= oCharsToKeepAfterProject Then
			oReferenceSuffix = oAfterProject.Substring(0, oCharsToKeepAfterProject)
		Else
			oReferenceSuffix = oAfterProject ' Use all available characters if less than needed
		End If
	Else
		' If it doesn't start with "Project", use a fallback approach
		' You might want to modify this based on your specific needs
		oReferenceSuffix = oOriginalName
	End If
	
	' Save with new naming convention
	If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
		oRefDoc.SaveAs(Left(oADoc.FullFileName, Len(oADoc.FullFileName) -4) & "\" & oNewPrefix & oReferenceSuffix & ".ipt", True)
	ElseIf oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
		oRefDoc.SaveAs(Left(oADoc.FullFileName, Len(oADoc.FullFileName) -4) & "\" & oNewPrefix & oReferenceSuffix & ".iam", True)
	End If
	oRefDoc.Close
Next

Dim oOccDoc As Document
Dim oOccNewFileName As String
For Each oOcc As ComponentOccurrence In oADoc.ComponentDefinition.Occurrences
	oOccDoc = oOcc.Definition.Document
	
	' Get original filename without extension
	Dim oOriginalName As String = IO.Path.GetFileNameWithoutExtension(oOccDoc.FullFileName)
	
	' Check if it starts with "Project" and extract the suffix
	If oOriginalName.StartsWith("Project") And oOriginalName.Length > 7 Then
		' Take the specified number of characters after "Project"
		Dim oAfterProject As String = oOriginalName.Substring(7)
		If oAfterProject.Length >= oCharsToKeepAfterProject Then
			oReferenceSuffix = oAfterProject.Substring(0, oCharsToKeepAfterProject)
		Else
			oReferenceSuffix = oAfterProject
		End If
	Else
		oReferenceSuffix = oOriginalName
	End If
	
	' Build new filename
	If oOccDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
		oOccNewFileName = Left(oADoc.FullFileName,Len(oADoc.FullFileName)-4) & "\" & oNewPrefix & oReferenceSuffix & ".ipt"
	ElseIf oOccDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
		oOccNewFileName = Left(oADoc.FullFileName,Len(oADoc.FullFileName)-4) & "\" & oNewPrefix & oReferenceSuffix & ".iam"
	End If
	oOcc.Replace(oOccNewFileName, True)
Next