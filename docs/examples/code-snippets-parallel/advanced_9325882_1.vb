' Title: ilogic rule to create a rectangular pattern in an assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-create-a-rectangular-pattern-in-an-assembly/td-p/9325882
' Category: advanced
' Scraped: 2025-10-09T09:00:57.610397

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
Dim oOccs As ComponentOccurrences = oADef.Occurrences
Dim oTO As TransientObjects = ThisApplication.TransientObjects
Dim oObjects As ObjectCollection = oTO.CreateObjectCollection
oObjects.Add(oOccs.ItemByName("P1"))
oObjects.Add(oOccs.ItemByName("P2"))
Dim oRowEntity As Object = oADef.WorkAxes.Item(3)
Dim oColumnEntity As Object = oADef.WorkAxes.Item(2)

Dim oPatt As RectangularOccurrencePattern
oPatt = oADef.OccurrencePatterns.AddRectangularPattern(oObjects, oColumnEntity, True, 0, 1, oRowEntity, True, 24, 3)
oPatt.Name = "Pattern 1"