' Title: what is the &quot;correct&quot; way to write this line of code for a rule?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/what-is-the-quot-correct-quot-way-to-write-this-line-of-code-for/td-p/13761867
' Category: advanced
' Scraped: 2025-10-07T14:03:40.989821

' This line doesn't work because Component.IsActive doesn't seem to accept a path to the Occurrence Pattern
' Component.IsActive(MakePath("iStairHandrail:1", "Array")) = BooleanParameter

' Set a reference to the Assembly Document
Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument

' Set a reference to the Assembly Component Definition
Dim oAsmCompDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition

' Set a reference to the Component Occurrences
Dim oOccurrences As ComponentOccurrences = oAsmCompDef.Occurrences

' Set a reference to a Component Occurrence
Dim oOcc As ComponentOccurrence = oOccurrences.ItemByName("iStairHandrail:1") ' You should probably change this name to avoid undesirable behavior

' Set a reference to the Occurrence Component Definition
Dim oOccCompDef As AssemblyComponentDefinition = oOcc.Definition

' Set a reference to the Occurrence Pattern
Dim oOccPattern As OccurrencePattern = oOccCompDef.OccurrencePatterns.Item("Array")

' Suppress or Unsuppress the Occurrence Pattern
If BooleanParameter = False Then
	oOccPattern.Suppress
Else
	oOccPattern.Unsuppress
End If