' Title: Switch active drawing Standard with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/switch-active-drawing-standard-with-ilogic/td-p/13772691#messageview_0
' Category: api
' Scraped: 2025-10-07T13:58:00.923503

Sub SetActiveDrawingStandardStyle(oDrawDoc As DrawingDocument, _
	Optional styleName As String = Nothing)
	Const sDefaultStyleName As String = "Default"
	If String.IsNullOrWhiteSpace(styleName) Then
		styleName = sDefaultStyleName
	End If
	Dim oSMgr As DrawingStylesManager = oDrawDoc.StylesManager
	Dim oStStyle As DrawingStandardStyle = Nothing
	Try
		oStStyle = oSMgr.StandardStyles.Item(styleName)
	Catch oEx As Exception
		Logger.Error("Failed to find specified DrawingStanardStyle to make active." _
		& vbCrLf & oEx.Message & vbCrLf & oEx.StackTrace)
	End Try
	If Not oStStyle Is Nothing Then
		If oStStyle.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
			oStStyle = oStStyle.ConvertToLocal()
		ElseIf oStStyle.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
			oStStyle.UpdateFromGlobal()
		End If
		oSMgr.ActiveStandardStyle = oStStyle
	End If
End Sub