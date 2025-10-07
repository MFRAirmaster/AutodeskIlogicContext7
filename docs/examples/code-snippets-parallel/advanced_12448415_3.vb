' Title: iLogic to Rename Browser Nodes to &quot;Default&quot; setting?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-rename-browser-nodes-to-quot-default-quot-setting/td-p/12448415#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:31:42.916569

For Each occ As ComponentOccurrence In ThisApplication.ActiveDocument.ComponentDefinition.Occurrences
	occ.Name = "Part Number"
Next