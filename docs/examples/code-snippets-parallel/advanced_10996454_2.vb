' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454
' Category: advanced
' Scraped: 2025-10-09T09:00:46.753538

Sub Main
	
Dim oAsm As AssemblyDocument= ThisApplication.ActiveDocument
On Error Resume Next

'get the active assembly
Dim oAsmComp As AssemblyComponentDefinition= ThisApplication.ActiveDocument.ComponentDefinition
For Each öConstraint In oAsmComp.Constraints

öConstraint.Delete
Next
'set the Master LOD active
Dim oLODRep As LevelOfDetailRepresentation
oLODRep = oAsmComp.RepresentationsManager.LevelOfDetailRepresentations.Item("Master")
oLODRep.Activate

'Iterate through all of the top level occurrences
Dim oOccurrence As ComponentOccurrence
For Each oOccurrence In oAsmComp.Occurrences

'	'ground everything in the top level
	oOccurrence.Grounded = True

Next


	Dim oUM As UnitsOfMeasure = oAsm.UnitsOfMeasure
	Dim oOcc As ComponentOccurrence

For Each oOcc In oAsm.ComponentDefinition.Occurrences

	Dim oTransform As Matrix
   	oTransform = oOcc.Transformation

	Dim oOriginLocation As Vector
	oOriginLocation = oTransform.Translation


Dim AbstandvonEbene(0 To 3) As Double

AbstandvonEbene(1) = oOriginLocation.X
AbstandvonEbene(2) = oOriginLocation.Y
AbstandvonEbene(3) = oOriginLocation.Z



	'Create a proxy for Face0 (The face in the context of the assembly)
Dim Zähler As Integer = 1

'cycle each Origin plane in the top assembly
For Zähler = 1 To 3

Dim curAsmOrPlane As WorkPlane = oAsmComp.WorkPlanes.Item(Zähler)
'oOcc.CreateGeometryProxy(curAsmOrPlane1, curAsmOrPlane)

'Dim oComp1 As ComponentOccurrence = oADef.Occurrences.Item(1)

Dim oCompAsmDef As AssemblyComponentDefinition
Dim oCompPtDef As PartComponentDefinition
'cycle each Origin plane in the first level Occurence in the assembly
For Zähler2 = 1 To 3
	
If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
	oCompAsmDef = oOcc.Definition
	oOcc.CreateGeometryProxy(oCompAsmDef.WorkPlanes.Item(Zähler2),curCompOriPlane)

ElseIf oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
	oCompPtDef = oOcc.Definition
	oOcc.CreateGeometryProxy(oCompPtDef.WorkPlanes.Item(Zähler2), curCompOriPlane)
	
	

	
End If

Dim oParalell As Boolean= curAsmOrPlane.Plane.IsParallelTo(curCompOriPlane.Plane, 0.00001)
Dim oSameDirection As Boolean= curAsmOrPlane.Plane.Normal.IsEqualTo(curCompOriPlane.Plane.Normal)

	Dim oNV As NameValueMap


	If oParalell=True Then
'

				If oSameDirection=True Then
				oConstraint=oAsmComp.Constraints.AddFlushConstraint(curAsmOrPlane, curCompOriPlane, AbstandvonEbene(Zähler))
				End If
				If oSameDirection=False Then
				oConstraint=oAsmComp.Constraints.AddMateConstraint(curAsmOrPlane, curCompOriPlane, AbstandvonEbene(Zähler))
				End If

	  End If

	Next
	

Next	

	oOcc.Grounded = False

Next
End Sub