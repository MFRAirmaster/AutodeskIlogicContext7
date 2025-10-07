' Title: External rule to run first local form of a part / assembly / drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/external-rule-to-run-first-local-form-of-a-part-assembly-drawing/td-p/13767020#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:38:48.754269

Dim oLocalFormNames As IEnumerable(Of String) = iLogicForm.FormNames
If oLocalFormNames IsNot Nothing AndAlso oLocalFormNames.Any() Then
	iLogicForm.Show(oLocalFormNames.First())
End If