' Title: Code doesn't work anymore for dxf
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/code-doesn-t-work-anymore-for-dxf/td-p/13763668#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:31:26.320363

Dim oInvApp As Inventor.Application = ThisApplication
Dim oDoc As Inventor.Document = oInvApp.ActiveDocument
If oDoc Is Nothing Then Return
If (Not TypeOf oDoc Is PartDocument) AndAlso (Not TypeOf oDoc Is AssemblyDocument) Then
	MsgBox("Active Document Is Not An Assembly or Part!", vbCritical, "")
	Return
End If
'<<< verify pre-selection >>>
Dim oSS As Inventor.SelectSet = oDoc.SelectSet
If oSS.Count = 0 Then
	MsgBox("Nothing Pre-Selected!", vbCritical, "")
	Return
ElseIf oSS.Count > 1 Then
	MsgBox("Multiple Objects Pre-Selected!", vbCritical, "")
	Return
End If
'try to Cast the first generic Object to a Face type (no error on failure)
Dim oFace As Inventor.Face = TryCast(oSS.Item(1), Inventor.Face)
'make sure the variable got assigned a value
If oFace Is Nothing Then
	MsgBox("No Face Pre-Selected!", vbCritical, "")
	Return
End If
'<<< now formulate the full path and file name of the new DXF file >>>
If oDoc.FileSaveCounter = 0 Then
	MsgBox("This Document has not been saved yet, so it has no 'Path' or 'FileName'!", vbCritical, "")
	Return
End If
Dim sFFN As String = oDoc.FullFileName
'gets the full path, without file name
Dim sPath As String = System.IO.Path.GetDirectoryName(sFFN)
'combine that path with sub folder name
sPath = System.IO.Path.Combine(sPath, "DXF")
'if that folder does not exist, create it
If Not System.IO.Directory.Exists(sPath) Then
	System.IO.Directory.CreateDirectory(sPath)
End If
'get just the file name, without path, and without file extension
Dim sFileName As String = System.IO.Path.GetFileNameWithoutExtension(sFFN)
'combine path and new file name for full file name of new DXF
sFFN = System.IO.Path.Combine(sPath, sFileName & ".dxf")
Logger.Info("Proposed FullFileName of new DXF:" & vbCrLf & sFFN)

Dim oCmdMgr As Inventor.CommandManager = oInvApp.CommandManager
'send data to Inventor's memory
oCmdMgr.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, sFFN)
'give Inventor time to process that task
oInvApp.UserInterfaceManager.DoEvents()
'capture Face Edges, and create DXF from them (normally shows SaveAs dialog)
oCmdMgr.ControlDefinitions.Item("GeomToDXFCommand").Execute()
'give Inventor time to process that task
oInvApp.UserInterfaceManager.DoEvents()
'open folder where file was created
Shell("explorer.exe " & sPath, vbNormalFocus)