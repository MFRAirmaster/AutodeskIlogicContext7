' Title: Excluding Print and Count if Sheet name matches
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697#messageview_0
' Category: api
' Scraped: 2025-10-09T08:59:35.652250

Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeVendorSheets As Boolean = InputRadioBox("Exclude Vendor Sheets?", "True", "False", False, "Toggle Vendor Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("Vendor") Then
		oSheet.ExcludeFromCount = oExcludeVendorSheets
		oSheet.ExcludeFromPrinting = oExcludeVendorSheets
	End If
Next