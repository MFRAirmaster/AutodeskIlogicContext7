' Title: External rule to run first local form of a part / assembly / drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/external-rule-to-run-first-local-form-of-a-part-assembly-drawing/td-p/13767020#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:08:47.292885

Dim oLocalFormNames As IEnumerable(Of String) = iLogicForm.FormNames
Dim oDoc As Document = ThisApplication.ActiveDocument
If oLocalFormNames.Count = 0 Then MessageBox.Show("'" & oDoc.DisplayName & "' doesn't provide a local form.", "No form available")
If oLocalFormNames.Count = 1 Then iLogicForm.Show(oLocalFormNames.First())
If oLocalFormNames.Count > 1 Then
	sSel = InputListBox("if nothing is selected the first form will be selected", iLogicForm.FormNames, iLogicForm.FormNames.First, "Choose Form", "Available Forms")
	iLogicForm.Show(sSel)
End If