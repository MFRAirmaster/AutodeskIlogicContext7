' Title: Suppress/make invisible weldment features with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/suppress-make-invisible-weldment-features-with-ilogic/td-p/13826974
' Category: advanced
' Scraped: 2025-10-09T08:51:17.203399

Dim oDoc As AssemblyDocument = ThisDoc.Document
For Each oOccurrence As ComponentOccurrence In oDoc.ComponentDefinition.Occurrences
	If (oOccurrence.Name.Contains("Weld"))
		For Each oSubOcc As ComponentOccurrence In oOccurrence.SubOccurrences
			If (oSubOcc.Name.Contains("Tray Weld")) Then
				oSubOcc.Visible = False
			End If
		Next
	End If
Next