' Title: is there a way to find all Browser notes that does not show the partnumber ?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/is-there-a-way-to-find-all-browser-notes-that-does-not-show-the/td-p/13815636
' Category: advanced
' Scraped: 2025-10-07T13:57:28.372632

Dim ass As AssemblyDocument = ThisDoc.Document
For Each occ As ComponentOccurrence In ass.ComponentDefinition.Occurrences
	If Split(occ.Name, ":")(0) <> occ.Definition.Document.propertysets(3)("Part Number").value Then 
	occ.Name = ""
	If Split(occ.Name, ":")(0) <> occ.Definition.Document.propertysets(3)("Part Number").value Then 
	Dim doc As PartDocument = occ.Definition.Document
	doc.DisplayName = doc.propertysets(3)("Part Number").value
	occ.Name = ""
	End If 
	End If 
Next