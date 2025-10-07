' Title: Changing sheetmetal style of a Lofted Flange using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/changing-sheetmetal-style-of-a-lofted-flange-using-ilogic/td-p/13824763#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:59:38.209153

Dim oPartDoc As PartDocument = ThisDoc.ModelDocument
Dim oSmDef As SheetMetalComponentDefinition=oPartDoc.ComponentDefinition
Dim oFeatures As PartFeatures = oSmDef.Features

For Each oFeature As PartFeature In oFeatures
	If oFeature.Name() ="Loft 1" Then
		Dim oLoftedFlangeDef As LoftedFlangeDefinition = oFeature.Definition()
		oLoftedFlangeDef.SheetMetalRule=oSmDef.SheetMetalStyles("1,0")
	Else If oFeature.Name() = "Loft 2" Then
		Dim oLoftedFlangeDef As LoftedFlangeDefinition = oFeature.Definition()
		oLoftedFlangeDef.SheetMetalRule=oSmDef.SheetMetalStyles("2,0")
	End If
Next