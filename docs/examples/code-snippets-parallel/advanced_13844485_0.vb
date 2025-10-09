' Title: Part rule with Ref parameters to top assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/part-rule-with-ref-parameters-to-top-assembly/td-p/13844485
' Category: advanced
' Scraped: 2025-10-09T09:04:54.577599

Dim oPDoc As PartDocument = ThisDoc.Document
Dim oPDef As PartComponentDefinition = oPDoc.ComponentDefinition
Dim oSketch As PlanarSketch = oPDef.Sketches.Item("TANK SHELL")
Dim oDims As DimensionConstraints = oSketch.DimensionConstraints
Dim oHeightDim As DimensionConstraint
Dim oAngleDim As DimensionConstraint
For Each oDim As DimensionConstraint In oDims
	If oDim.Parameter.Name = "TANK_EXTERNAL_DIA" Then
		oHeightDim = oDim
	ElseIf oDim.Parameter.Name = "TANK_INTERNAL_DIA" Then
		oAngleDim = oDim
	End If
Next

TANK_EXTERNAL = Not TANK_INTERNAL

Try
	oAngleDim.Driven = Not oAngleDim.Driven
Catch
	oHeightDim.Driven = Not oHeightDim.Driven
	oAngleDim.Driven = Not oAngleDim.Driven
	Exit Sub
End Try
oHeightDim.Driven = Not oHeightDim.Driven