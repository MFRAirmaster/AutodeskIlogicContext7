' Title: printing Drawings in A0,A1,A2 on our Canon Plotter
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/printing-drawings-in-a0-a1-a2-on-our-canon-plotter/td-p/13832687
' Category: basic
' Scraped: 2025-10-07T14:06:55.552412

printoptions = {"Current Sheet - Full Size", "All Sheets - Full Size", "Sheet Range - Full Size", "Current Sheet - A Size", "All Sheets - A Size", "Sheet Range - A Size"}
printoptions = InputListBox("Choose printing option.", MultiValue.List("printoptions"), printoptions, Title := "Printing Options", ListName := "Print Options")

Select printoptions
	Case "Current Sheet - Full Size"
		oSheet = ThisDrawing.ActiveSheet.Sheet
		Call oPrint
		
	Case "All Sheets - Full Size"
		oSheets = oDoc.Sheets
		For Each oSheet In oSheets
			Call oPrint
		Next
'add in the rest
End Select