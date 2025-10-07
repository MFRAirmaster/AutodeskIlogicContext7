' Title: Create work point from axis and plane
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-work-point-from-axis-and-plane/td-p/11667060#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:05:16.064711

Sub CreateWorkPoint()

Dim oPartDoc As PartDocument
Set oPartDoc = ThisApplication.ActiveDocument

Dim oCompDef As ComponentDefinition
Set oCompDef = oPartDoc.ComponentDefinition

Dim face As Object
Set face = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAllPlanarEntities, "Select Face")
Dim axis As Object
Set axis = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAllPlanarEntities, "Select Axis")

Call oCompDef.WorkPoints.Addby <---- I'm stuck here

End Sub