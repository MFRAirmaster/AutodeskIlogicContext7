' Title: Create work point from axis and plane
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-work-point-from-axis-and-plane/td-p/11667060
' Category: advanced
' Scraped: 2025-10-09T09:00:48.735295

Sub Main
	'pick face
	oObj = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Pick Flat Face.")
	If oObj Is Nothing OrElse (TypeOf oObj Is Face = False) Then Return
	Dim oFace As Face = oObj
	Dim oPlane As Plane = oFace.Geometry
	
	'pick work axis
	oObj = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkAxisFilter, "Pick Work Axis.")
	If oObj Is Nothing OrElse (TypeOf oObj Is WorkAxis = False) Then Return
	Dim oAxis As WorkAxis = oObj
	
	Dim oPoint As Point = oPlane.IntersectWithLine(oAxis.Line)

	Dim oPDef As PartComponentDefinition = oFace.Parent.ComponentDefinition
	oPDef.WorkPoints.AddByPoint(oPoint)
End Sub