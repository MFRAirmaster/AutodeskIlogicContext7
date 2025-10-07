' Title: ilogic rule to create a rectangular pattern in an assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-create-a-rectangular-pattern-in-an-assembly/td-p/9325882
' Category: advanced
' Scraped: 2025-10-07T14:03:29.616764

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
Dim oOccs As ComponentOccurrences = oADef.Occurrences
Dim oOcc1 As ComponentOccurrence = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter,"Select the first Occurrence to pattern.")
Dim oOcc2 As ComponentOccurrence = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter,"Select the second Occurrence to pattern.")

Dim oTO As TransientObjects = ThisApplication.TransientObjects
Dim oObjects As ObjectCollection = oTO.CreateObjectCollection
oObjects.Add(oOcc1)
oObjects.Add(oOcc2)
'oObjects.Add(oOccs.Item(1))
'oObjects.Add(oOccs.Item(2))
Dim oRowEntity As Object = oADef.WorkAxes.Item(1)
Dim oColumnEntity As Object = oADef.WorkAxes.Item(2)

Dim oPatt As RectangularOccurrencePattern
oPatt = oADef.OccurrencePatterns.AddRectangularPattern(oObjects, oColumnEntity, True, 0, 1, oRowEntity, True, Parameter("Distance"), Parameter("Quantity"))
MsgBox("Check Model screen to make sure they are patterned right." & vbNewLine &
"Because, after you click OK, they will be deleted again.")
'Do whatever other stuff here

'For i=1 To oPatt.ParentComponents.Count
'	oPatt.ParentComponents.Remove(i)
'Next
oPatt.Delete
oOcc1.Delete
oOcc2.Delete