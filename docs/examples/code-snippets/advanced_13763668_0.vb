' Title: Code doesn't work anymore for dxf
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/code-doesn-t-work-anymore-for-dxf/td-p/13763668#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:31:26.320363

'Export face to pre-defined folder

If ThisApplication.ActiveDocument.SelectSet.Count = 0 Then
	MsgBox("Face not selected. Aborting Rule!")
	Exit Sub
End If

oPath = ThisDoc.Path
'get DXF target folder path
oFolder = oPath & "\" & "DXF"
'Check for the dxf folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
System.IO.Directory.CreateDirectory(oFolder)
End If
oFileName = oFolder & "\" & ThisDoc.FileName & ".dxf"

Dim oCmdMgr As CommandManager
oCmdMgr = ThisApplication.CommandManager

Call oCmdMgr.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, oFileName)
Call oCmdMgr.ControlDefinitions.Item("GeomToDXFCommand").Execute

'open the folder where the new files are saved
Shell("explorer.exe " & oFolder,vbNormalFocus)