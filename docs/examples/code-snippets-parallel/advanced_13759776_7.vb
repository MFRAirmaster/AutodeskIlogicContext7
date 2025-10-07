' Title: Defining A-Side for sheet metal part using Entitys
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/defining-a-side-for-sheet-metal-part-using-entitys/td-p/13759776#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:10:42.569110

Dim oPartDoc As PartDocument = ThisApplication.ActiveDocument
Dim oPartCompDef As SheetMetalComponentDefinition = oPartDoc.ComponentDefinition

Dim bHasFlatPattern As Boolean = oPartCompDef.HasFlatPattern
If bHasFlatPattern = True Then oPartCompDef.FlatPattern.Delete

Dim oASideDefinitions As ASideDefinitions = oPartCompDef.ASideDefinitions

If oASideDefinitions.Count <> 0 Then
	For i As Integer = 1 To oASideDefinitions.Count
		oASideDefinitions.Item(i).Delete
	Next
End If
Dim A_Side
If Orientation = "RD"
	A_Side = "RD_Back"
Else 
	A_Side = "LG_Back"
End If


Dim sNamedEntityName As String = A_Side
Dim oFace As Face = ThisDoc.NamedEntities.TryGetEntity(sNamedEntityName)

If oFace IsNot Nothing Then oASideDefinitions.Add(oFace)

If bHasFlatPattern = True Then oPartCompDef.Unfold