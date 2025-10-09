' Title: External rule to run first local form of a part / assembly / drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/external-rule-to-run-first-local-form-of-a-part-assembly-drawing/td-p/13767020#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:08:47.292885

Dim oLocalFormNames As IEnumerable(Of String) = iLogicForm.FormNames
If oLocalFormNames.Count = 0 Then MessageBox.Show("Document doesn't provide a local form.", "No form available")
If oLocalFormNames.Count = 1 Then iLogicForm.Show(oLocalFormNames.First())
If oLocalFormNames.Count > 1 Then
	Dim oAllFormNames As New ArrayList
	For Each Name In iLogicForm.FormNames
		oAllFormNames.Add(Name)
	Next
	sSel = InputListBox("if nothing is selected the first form will be selected", oAllFormNames, (1), "Choose Form", "Available Forms")
	If sSel = "" Then
		iLogicForm.Show(oLocalFormNames.First())
	Else
		iLogicForm.Show(sSel)
	End If
End If