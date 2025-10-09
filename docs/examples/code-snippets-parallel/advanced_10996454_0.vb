' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454
' Category: advanced
' Scraped: 2025-10-09T09:00:46.753538

Sub Main
	
Dim oAsm As AssemblyDocument= ThisApplication.ActiveDocument
On Error Resume Next

'get the active assembly
Dim oAsmComp As AssemblyComponentDefinition= ThisApplication.ActiveDocument.ComponentDefinition
For Each oConstraint In oAsmComp.Constraints

oConstraint.Delete
Next
'set the Master LOD active
Dim oLODRep As LevelOfDetailRepresentation
oLODRep = oAsmComp.RepresentationsManager.LevelOfDetailRepresentations.Item("Master")
oLODRep.Activate

'Iterate through all of the top level occurrences
Dim oOccurrence As ComponentOccurrence
For Each oOccurrence In oAsmComp.Occurrences
'Try

'	'ground everything in the top level
	oOccurrence.Grounded = True
'Catch
'	end try
Next


	Dim oUM As UnitsOfMeasure = oAsm.UnitsOfMeasure
	Dim oOcc As ComponentOccurrence

	For Each oOcc In oAsm.ComponentDefinition.Occurrences


	'Create a proxy for Face0 (The face in the context of the assembly)
Dim Zähler As Integer = 1

'cycle each Origin plane in the top assembly
For Zähler = 1 To 3

Dim curAsmorPlane As WorkPlane = oAsmComp.WorkPlanes.Item(Zähler)

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
	oOcc.CreateGeometryProxy(oCompPtDef.WorkPlanes.Item(Zähler2),curCompOriPlane)
	
End If

	'Measure distance
	Dim oNV As NameValueMap
'	Dim Distance As Double

'	Try
'Dim angle As Double = ThisApplication.MeasureTools.
	Dim angle As Double = ThisApplication.MeasureTools.GetAngle(curAsmorPlane, curCompOriPlane)
	'Angle in deg 
'	MsgBox(angle)
	angle = (angle * 180) / PI

	If angle=0 Then
'
				'Convert distance from database units to default units of the document
				Dim Distance1 As Double = ThisApplication.MeasureTools.GetMinimumDistance(curAsmorPlane, curCompOriPlane)
				Dim Distance As Double 
				Dim oConstraint As AssemblyConstraint
				Distance = oUM.ConvertUnits(Distance1, UnitsTypeEnum.kDatabaseLengthUnits, oUM.LengthUnits)
				'Return the value in a messagebox just to control that it's right
				Distance= Round(Distance,3)
'				MsgBox(Distance)
				
				oConstraint=oAsmComp.Constraints.AddMateConstraint(curAsmorPlane, curCompOriPlane, Distance1)
				
		      If oConstraint.HealthStatus = oConstraint.HealthStatus.kInconsistentHealth Then
	          oConstraint.Delete
			  oConstraint = oAsmComp.Constraints.AddMateConstraint(curAsmorPlane, curCompOriPlane, -Distance1)
			  End If
		      If oConstraint.HealthStatus = oConstraint.HealthStatus.kInconsistentHealth Then
	          oConstraint.Delete
			  oConstraint = oAsmComp.Constraints.AddFlushConstraint(curAsmorPlane, curCompOriPlane, Distance1)
			  End If
		      If oConstraint.HealthStatus = oConstraint.HealthStatus.kInconsistentHealth Then
	          oConstraint.Delete
			  oConstraint = oAsmComp.Constraints.AddFlushConstraint(curAsmorPlane, curCompOriPlane, -Distance1)
			  End If
		      If oConstraint.HealthStatus = oConstraint.HealthStatus.kInconsistentHealth Then
	          oConstraint.Delete
			  End If
	  End If
'	Catch
'	End Try
	Next
	

Next	
'Try 
	oOcc.Grounded = False
'	Catch 
'		End try
Next
End Sub