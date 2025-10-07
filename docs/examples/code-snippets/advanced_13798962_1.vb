' Title: Turn off visibility of welds in an occurrence
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/turn-off-visibility-of-welds-in-an-occurrence/td-p/13798962#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:35:45.764362

For Each oOccurrence As ComponentOccurrence In oAssy.ComponentDefinition.Occurrences
	If (oOccurrence.Name.Contains("Weld"))
		For Each oSubOcc As ComponentOccurrence In oOccurrence.SubOccurrences
			oSubOcc.Visible = False
		Next
	End If
Next