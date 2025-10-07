' Title: search assembly containing virtual component
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/search-assembly-containing-virtual-component/td-p/13812753
' Category: advanced
' Scraped: 2025-10-07T13:08:26.460516

Dim ass As AssemblyDocument = ThisDoc.Document
For Each occ As ComponentOccurrence In ass.ComponentDefinition.Occurrences
	If TypeOf occ.Definition Is VirtualComponentDefinition Then 
				MsgBox (virtual)
	End If
Next