' Title: iLogic to Rename Browser Nodes to &quot;Default&quot; setting?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-rename-browser-nodes-to-quot-default-quot-setting/td-p/12448415#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:31:42.916569

Dim oADoc As AssemblyDocument = TryCast(ThisDoc.Document, Inventor.AssemblyDocument)
If oADoc Is Nothing Then Return
For Each oOcc As ComponentOccurrence In oADoc.ComponentDefinition.Occurrences
	Dim sPN As String = iProperties.Value(oOcc.Name, "Project", "Part Number")
	Try
		oOcc.Name = sPN
	Catch
		Logger.Error("Could not rename component named: " & oOcc.Name _
		& vbCrLf & "to: " & sPN)
	End Try
Next
oADoc.Update2(True)