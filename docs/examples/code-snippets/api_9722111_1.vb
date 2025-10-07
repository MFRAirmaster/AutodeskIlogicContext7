' Title: Add leading zeros to title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111#messageview_0
' Category: api
' Scraped: 2025-10-07T13:29:45.677126

'Get Total Number of Sheets
SheetCount = ThisApplication.ActiveDocument.Sheets.Count

'Append a leading zero if less than 10
If SheetCount < 10 Then
	SheetCo = "0" & SheetCount
Else 
	SheetCo = SheetCount
End If

'Loop through each sheet and grab the sheet number from the name
For Each Sheet In ThisApplication.ActiveDocument.Sheets
SheetNumber = Mid(Sheet.Name, InStr(1, Sheet.Name, ":") + 1)

'Append a leading zero if less than 10
If SheetNumber <10 Then 
SheetNo = "0" & SheetNumber
Else
	SheetNo = SheetNumber
End If

MessageBox.Show(SheetCo & SheetNo, "Sheet descriptor")

Next