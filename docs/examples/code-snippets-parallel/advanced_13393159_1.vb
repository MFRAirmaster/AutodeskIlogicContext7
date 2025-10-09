' Title: CUSTOM VIEW ORIENTATION ALTERNATIVE
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-view-orientation-alternative/td-p/13393159#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:54:01.833158

Public Sub Main()
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oDDoc As DrawingDocument = oInvApp.ActiveDocument
	If oDDoc Is Nothing Then Exit Sub
	Dim oView As DrawingView
	oView = oInvApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select View...")
	If oView Is Nothing Then Exit Sub
	If Not oView.ViewStyle = kShadedDrawingViewStyle Then
		Call oDDoc.SelectSet.Select(oView)
		call oInvApp.CommandManager.ControlDefinitions("DrawingUpdateViewComponentCmd").Execute
	End If
	Dim oDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument
	Dim oPlane As WorkPlane = TryCast(oDoc.ComponentDefinition.WorkPlanes("FRONT_VIEW_PLANE"), WorkPlane)
	If oPlane Is Nothing Then Exit Sub
    Dim oNormalEye As Inventor.Vector = oPlane.Plane.Normal.AsVector
    Dim oPosPoint As Point = Nothing
    Dim oPosVectorX As UnitVector = Nothing
    Dim oPosVectorY As UnitVector = Nothing
	oPlane.GetPosition(oPosPoint, oPosVectorX, oPosVectorY)	
	With oView.Camera
		.Target = oPosPoint
		oPosPoint.TranslateBy(oNormalEye)
	    .Eye = oPosPoint
	    .UpVector = oPosVectorY
		.ApplyWithoutTransition()
	End With	
End Sub