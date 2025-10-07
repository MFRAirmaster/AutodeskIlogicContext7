' Title: Trying to detect if the All  Workfeatures button is checked or not.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trying-to-detect-if-the-all-workfeatures-button-is-checked-or/td-p/13766498#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:21:39.014817

Sub Main()
	
	Dim invApp As Inventor.Application = ThisApplication
	Dim oControlDefs As ControlDefinitions
	Dim oControlDef As ControlDefinition
	oControlDefs = invApp.CommandManager.ControlDefinitions
	
	For Each oControlDef In oControlDefs
		
		If oControlDef.DisplayName = "All Workfeatures" Then
			'MessageBox.Show("Found it")
			oControlDef.Execute
		end if
		
	Next
	
End Sub