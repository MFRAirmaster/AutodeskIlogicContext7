' Title: Save and replace parts in an assembly using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/save-and-replace-parts-in-an-assembly-using-ilogic/td-p/9409910#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:58:15.609350

On Error Resume Next
oSelect = ThisDoc.Document.SelectSet.Item(1)

'catch an empty string in the imput
If Err.Number <> 0 Then
MessageBox.Show("Please select a component before running this rule.", "iLogic")
Return
End If

' get a component by name.
Dim compOcc = Component.InventorComponent(oSelect.name) 


'define document
oDoc = compOcc.Definition.Document
'create a file dialog box
Dim oFileDlg As Inventor.FileDialog = Nothing
InventorVb.Application.CreateFileDialog(oFileDlg)

'check file type and set dialog filter
If oDoc.DocumentType = kPartDocumentObject Then
oFileDlg.Filter = "Autodesk Inventor Part Files (*.ipt)|*.ipt"
Else If oDoc.DocumentType = kAssemblyDocumentObject Then
oFileDlg.Filter = "Autodesk Inventor Assembly Files (*.iam)|*.iam"
End If

'set the directory to open the dialog at
oFileDlg.InitialDirectory = ThisDoc.WorkspacePath()
'set the file name string to use in the input box
oFileDlg.FileName = System.IO.Path.GetFileName(oDoc.FullFileName)

'work with an error created by the user backing out of the save
oFileDlg.CancelError = True
On Error Resume Next
'specify the file dialog as a save dialog (rather than a open dialog)
oFileDlg.ShowSave()

'catch an empty string in the imput
If Err.Number <> 0 Then
MessageBox.Show("No File Saved.", "iLogic: Dialog Canceled")
ElseIf oFileDlg.FileName <> "" Then
MyFile = oFileDlg.FileName
'save the file
oDoc.SaveAs(MyFile, False) 'True = Save As Copy & False = Save As
End If