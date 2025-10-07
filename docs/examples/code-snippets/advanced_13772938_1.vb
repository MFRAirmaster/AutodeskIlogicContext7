' Title: Mark feature messes up loop count in iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/mark-feature-messes-up-loop-count-in-ilogic/td-p/13772938
' Category: advanced
' Scraped: 2025-10-07T13:37:01.363013

Sub Main
	
	Dim isdoc As Document = ThisDoc.Document
	If isdoc.SubType <> "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
		Exit Sub
	Else
		'Create custom iproperties
		customPropertySet = ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")
		
		Dim iprop(4) As String
		iprop(1) = "Cutting Length"
		iprop(2) = "Pierces"
		iprop(3) = "Bends"
		iprop(4) = "Weight"
		
		For k = 1 To 4
		Dim prop(k) As String
		    Try
		        prop(k) = iProperties.Value("Custom", iprop(k))
		    Catch
		        'Assume error means not found
		        customPropertySet.Add("", iprop(k))
		        iProperties.Value("Custom", iprop(k)) = "null"
		    End Try
		Next
		
		'Get cutting length
		Dim oSMDoc As PartDocument = ThisApplication.ActiveDocument
		Dim oSMDDef As SheetMetalComponentDefinition = oSMDoc.ComponentDefinition
		If oSMDDef.HasFlatPattern = False Then
			MsgBox("There is no Flat Pattern yet.", vbOKOnly, "NO FLAT PATTERN")
			Return
		End If
		Dim oTopFace As Face = oSMDDef.FlatPattern.TopFace
		Dim oMeasureTools As MeasureTools = ThisApplication.MeasureTools
		Dim oLoopLength As Double = 0
		For Each oEdgeLoop As EdgeLoop In oTopFace.EdgeLoops
			oLoopLength = oLoopLength + oMeasureTools.GetLoopLength(oEdgeLoop)
		Next
		oLoopLength = oLoopLength / 2.54
		oLoopLength = Round(oLoopLength, 2)

		'get bends
		Dim bendqty As Integer = oSMDDef.Bends.Count
		
		'get mass
		mass = Round(iProperties.Mass, 2)
	
		'get pierces
		TotalPierces = 1
		For Each oEdgeLoop In oSMDDef.FlatPattern.TopFace.EdgeLoops
		    If oEdgeLoop.IsOuterEdgeLoop = False Then
		        TotalPierces = TotalPierces + 1
		    End If
		Next
		
		'write data to iproperties
		iProperties.Value("Custom", "Pierces") = TotalPierces
		iProperties.Value("Custom", "Bends") = bendqty
		iProperties.Value("Custom", "Cutting Length") = oLoopLength
		iProperties.Value("Custom", "Weight") = mass & " LBS."
		
	End If
	End Sub