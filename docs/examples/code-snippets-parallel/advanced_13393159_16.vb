' Title: CUSTOM VIEW ORIENTATION ALTERNATIVE
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-view-orientation-alternative/td-p/13393159
' Category: advanced
' Scraped: 2025-10-07T14:07:54.930591

Public Sub Main()
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oPDoc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
	If oPDoc Is Nothing Then Exit Sub
	Dim oPDef As PartComponentDefinition = oPDoc.ComponentDefinition
	Dim oPlane As WorkPlane = TryCast(oPDef.WorkPlanes(sNamePlane), WorkPlane)
	If oPlane Is Nothing Then Exit Sub
	If oPDoc.Dirty Then oPDoc.Update()
	Dim oTG As Transaction = oInvApp.TransactionManager.StartTransaction(oPDoc, "Set Camera...")
	Dim oNewView As DesignViewRepresentation
	Dim oActView As DesignViewRepresentation
	With oPDef.RepresentationsManager
		oActView = .ActiveDesignViewRepresentation
		If oActView.Name = sNameView Then
			oNewView = oActView
		Else
			Try : oNewView = .DesignViewRepresentations(sNameView)
			Catch : oNewView = .DesignViewRepresentations.Add(sNameView)
			End Try
		End If
	End With
	oNewView.Activate()
	Call SetCamera(oNewView.Camera, oPlane)
	Try : oActView.Activate() : Catch : End Try
	oTG.End()
End Sub

Property sNamePlane As String = "FRONT_VIEW_PLANE"
Property sNameView As String = "DWG_FRONT"

Private Function SetCamera(ByVal oCamera As Camera, ByVal oPlane As WorkPlane)
	Dim oNormalEye As Inventor.Vector = oPlane.Plane.Normal.AsVector()
    Dim oPosPoint As Point = Nothing
    Dim oPosVectorX As UnitVector = Nothing
    Dim oPosVectorY As UnitVector = Nothing
	oPlane.GetPosition(oPosPoint, oPosVectorX, oPosVectorY)
	oCamera.Target = oPosPoint
	oPosPoint.TranslateBy(oNormalEye)
    oCamera.Eye = oPosPoint
'	With oPosVectorY
'		Dim dCoors() As Double = {-.X, -.Y, -.Z }
'		.PutUnitVectorData(dCoors)	
'	End With
	oCamera.UpVector = oPosVectorY
	oCamera.ApplyWithoutTransition()
End Function