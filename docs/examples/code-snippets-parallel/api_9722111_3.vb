' Title: Add leading zeros to title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111#messageview_0
' Category: api
' Scraped: 2025-10-09T08:56:09.423332

Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)).ToString("00")
	MsgBox(oSheets.Count.ToString("00") & oSheetNumber)
Next