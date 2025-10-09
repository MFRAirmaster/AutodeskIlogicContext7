' Title: Help on API AddByiMateAndEntity
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/help-on-api-addbyimateandentity/td-p/13822199
' Category: advanced
' Scraped: 2025-10-09T08:50:31.486024

Dim oImate As Object

Dim oMateName As iMateDefinition

Dim oFace As face

....

Set oImate = oCompDef.iMateResults.AddByiMateAndEntity(oMateName.Name, oFace)