' Title: Excluding Print and Count if Sheet name matches
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697#messageview_0
' Category: api
' Scraped: 2025-10-09T08:59:35.652250

Dim oDoc As DrawingDocument
oDoc = ThisApplication.ActiveDocument
Dim oSheet As Sheet

i = 0
x = 0

For Each oSheet In oDoc.Sheets
	i = i + 1
Next

While (x<i)
	Sheet_Name = "Vendor:" & x 
	x= x+1
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromPrinting = True  <---Not working
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromCount = True <---Not working
End While