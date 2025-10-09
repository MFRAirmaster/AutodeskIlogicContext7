' Title: Set a component in an assembly with it's active view representation
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/set-a-component-in-an-assembly-with-it-s-active-view/td-p/13820679#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:59:26.276864

Dim oPickedOcc As ComponentOccurrence = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select a Component.")
If oPickedOcc Is Nothing Then Return
Dim oCD As Inventor.ComponentDefinition = oPickedOcc.Definition
Dim oRepsMgr As Inventor.RepresentationsManager = oCD.RepresentationsManager
Dim sDVR As String = oRepsMgr.ActiveDesignViewRepresentation.Name
oPickedOcc.SetDesignViewRepresentation(sDVR,,True)