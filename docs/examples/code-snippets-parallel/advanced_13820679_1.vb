' Title: Set a component in an assembly with it's active view representation
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/set-a-component-in-an-assembly-with-it-s-active-view/td-p/13820679
' Category: advanced
' Scraped: 2025-10-07T14:06:46.929115

Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument


Dim oOccurrences As ComponentOccurrences = oAsmDoc.ComponentDefinition.Occurrences

For Each oOccurrence As ComponentOccurrences In oOccurrences

	Dim oRepsMgr As Inventor.RepresentationsManager = oOccurrences.RepresentationsManager
	Dim sDVR As String = oRepsMgr.ActiveDesignViewRepresentation.Name
	oOccurrences.SetDesignViewRepresentation(sDVR,,True)
	
Next