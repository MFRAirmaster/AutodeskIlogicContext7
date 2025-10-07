' Title: Defining A-Side for sheet metal part using Entitys
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/defining-a-side-for-sheet-metal-part-using-entitys/td-p/13759776#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:57:19.073360

Dim partDoc As PartDocument = ThisDoc.Document
Dim compDef As SheetMetalComponentDefinition = partDoc.ComponentDefinition

If (compDef.HasFlatPattern) Then
	compDef.FlatPattern.Delete
End If

If Orientation = "RD"
	'compDef.ASideDefinitions.Add(RD Back) '<-- Not Working
Else If Orientation = "LG"
	'compDef.ASideDefinitions.Add(LG Back) '<-- Not Working
End If