' Title: How to &quot;Show All&quot; in a part document with iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-show-all-quot-in-a-part-document-with-ilogic/td-p/13832701#messageview_0
' Category: api
' Scraped: 2025-10-09T09:08:08.946645

Dim oApp As Inventor.Application = ThisApplication
Dim oPDoc As PartDocument = oApp.ActiveDocument
Dim oSBColl As ObjectCollection = oApp.TransientObjects.CreateObjectCollection()
For Each oSB As SurfaceBody In oPDoc.ComponentDefinition.SurfaceBodies
    oSBColl.Add(oSB)
Next
oPDoc.SelectSet.Clear()
oPDoc.SelectSet.SelectMultiple(oSBColl)
oApp.CommandManager.ControlDefinitions.Item("PartVisibilityCtxCmd").Execute()