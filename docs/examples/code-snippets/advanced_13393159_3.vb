' Title: CUSTOM VIEW ORIENTATION ALTERNATIVE
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-view-orientation-alternative/td-p/13393159
' Category: advanced
' Scraped: 2025-10-07T13:06:49.323707

Public Sub Main()
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oPDoc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
	If oPDoc Is Nothing Then Exit Sub
	Dim oPlane As WorkPlane = TryCast(oPDoc.ComponentDefinition.WorkPlanes("FRONT_VIEW_PLANE"), WorkPlane)
	If oPlane Is Nothing Then Exit Sub
	Dim oTG As Transaction = oInvApp.TransactionManager.StartTransaction(oPDoc, "Set Camera...")
	Dim oNormalEye As Inventor.Vector = oPlane.Plane.Normal.AsVector()
    Dim oPosPoint As Point = Nothing
    Dim oPosVectorX As UnitVector = Nothing
    Dim oPosVectorY As UnitVector = Nothing
	oPlane.GetPosition(oPosPoint, oPosVectorX, oPosVectorY)
	With oInvApp.ActiveView.Camera
		.Target = oPosPoint
		oPosPoint.TranslateBy(oNormalEye)
	    .Eye = oPosPoint
	    .UpVector = oPosVectorY
		.ApplyWithoutTransition()
		.Parent.SetCurrentAsFront()
	End With
	oTG.End()
End Sub