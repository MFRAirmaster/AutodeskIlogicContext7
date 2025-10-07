' Title: External rule to run first local form of a part / assembly / drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/external-rule-to-run-first-local-form-of-a-part-assembly-drawing/td-p/13767020#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:38:48.754269

Dim PossibleFormNames As New ArrayList				' list with all possible names
PossibleFormNames.Add("GO")
PossibleFormNames.Add("Values")
PossibleFormNames.Add("WebInput")					' Customer 1
PossibleFormNames.Add("Formular 1")					' Customer 2
PossibleFormNames.Add("FactoryDesignProperties")	' Factory 
PossibleFormNames.Add("DrawingControl")
PossibleFormNames.Add("nn")

Check = 0														' for debugging

For Each FormName In PossibleFormNames							' running through list
	Try
		If Check > 0 Then MsgBox("Try show form " & FormName)	' debugging
		iLogicForm.Show(FormName)								' try opening the form
		Exit Sub 												' if successful, leaving this rule
	Catch : End Try
Next

' feedback, in case of no form
MessageBox.Show("Document doesn't provide a local form.", "No Form available")