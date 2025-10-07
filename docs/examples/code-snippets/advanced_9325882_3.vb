' Title: ilogic rule to create a rectangular pattern in an assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-create-a-rectangular-pattern-in-an-assembly/td-p/9325882#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:21:30.983568

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
	Dim oOccs As ComponentOccurrences = oADef.Occurrences
	Dim oTO As TransientObjects = ThisApplication.TransientObjects
	Dim oObjects As ObjectCollection = oTO.CreateObjectCollection
	oObjects.Add(oOccs.ItemByName("4349:1"))
	oObjects.Add(oOccs.ItemByName("0157:1"))
	Dim oRowEntity As Object = oADef.WorkAxes.Item(3)
	Dim oColumnEntity As Object = oADef.WorkAxes.Item(2)

	Dim oPatt As RectangularOccurrencePattern
	oPatt = oADef.OccurrencePatterns.AddRectangularPattern(oObjects, oColumnEntity, True, 0, 1, oRowEntity, True, Parameter("DIST_SUP"), Parameter("QTD_SUP"))
	oPatt.Name = "PADR_PARAFUSOS_INOX"