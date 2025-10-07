' Title: Set a component in an assembly with it's active view representation
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/set-a-component-in-an-assembly-with-it-s-active-view/td-p/13820679#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:09:58.630323

Dim oADoc As AssemblyDocument = TryCast(ThisDoc.Document, Inventor.AssemblyDocument)
If oADoc Is Nothing Then Return
For Each oOcc As ComponentOccurrence In oADoc.ComponentDefinition.Occurrences
	If oOcc.Suppressed Then Continue For
	If TypeOf oOcc.Definition Is VirtualComponentDefinition Then Continue For
	If TypeOf oOcc.Definition Is WeldsComponentDefinition Then Continue For
	Dim sDVR As String = oOcc.Definition.RepresentationsManager.ActiveDesignViewRepresentation.Name
	oOcc.SetDesignViewRepresentation(sDVR,,True)
Next
oADoc.Update2(True)