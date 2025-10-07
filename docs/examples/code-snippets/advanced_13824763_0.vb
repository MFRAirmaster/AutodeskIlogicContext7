' Title: Changing sheetmetal style of a Lofted Flange using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/changing-sheetmetal-style-of-a-lofted-flange-using-ilogic/td-p/13824763#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:59:38.209153

Dim oPartDoc As PartDocument = ThisDoc.ModelDocument
Dim oFeatures As PartFeatures = oPartDoc.ComponentDefinition.Features

For Each oFeature As PartFeature In oFeatures
	If oFeature.Name() = "Loft 1" Then
		Dim oLoftedFlangeDef As LoftedFlangeDefinition = oFeature.Definition()
		oLoftedFlangeDef.SheetMetalRule.Thickness = LOFT_1_THICKNESS
	Else If oFeature.Name() = "Loft 2" Then
		Dim oLoftedFlangeDef As LoftedFlangeDefinition = oFeature.Definition()
		oLoftedFlangeDef.SheetMetalRule.Thickness = LOFT_2_THICKNESS
	End If
Next