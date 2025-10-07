' Title: ilogic rule to create a rectangular pattern in an assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-create-a-rectangular-pattern-in-an-assembly/td-p/9325882
' Category: advanced
' Scraped: 2025-10-07T14:03:29.616764

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
Dim oOccs As ComponentOccurrences = oADef.Occurrences
Dim oPatt1 As RectangularOccurrencePattern = oADef.OccurrencePatterns.Item("Pattern 1")
oPatt1.Delete
Dim oOcc1 As ComponentOccurrence = oOccs.ItemByName("P1")
Dim oOcc2 As ComponentOccurrence = oOccs.ItemByName("P2")
oOcc1.Delete
oOcc2.Delete