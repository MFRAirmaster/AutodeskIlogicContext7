' Title: iLogic to quickly add PDF as PNG background into current activ drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-quickly-add-pdf-as-png-background-into-current-activ/td-p/13757564#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:01:10.411531

Sub main()
	
	Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument
	Dim oSheet4 As Sheet = oDoc.Sheets(4)
	
	'This rule checks to see whether the user is on the correct page for placement before allowing the macro to bring in excel data. 

	If oDoc.ActiveSheet Is oSheet4 Then
		InventorVb.RunMacro("ApplicationProject", "Import_Excel_Module" , "Import_Excel")
	Else
		MessageBox.Show("Please make sure you have page 4 active and selected before running this rule.")
	End If 


End Sub