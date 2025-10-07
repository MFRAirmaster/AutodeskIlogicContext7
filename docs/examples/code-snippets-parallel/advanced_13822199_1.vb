' Title: Help on API AddByiMateAndEntity
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/help-on-api-addbyimateandentity/td-p/13822199
' Category: advanced
' Scraped: 2025-10-07T14:29:59.666111

Dim oImate As Object

Dim oMateName As iMateDefinition

Dim oFace As face

....

Set oImate = oCompDef.iMateResults.AddByiMateAndEntity(oMateName.Name, oFace)