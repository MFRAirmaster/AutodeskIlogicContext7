' Title: Add leading zeros to title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111#messageview_0
' Category: api
' Scraped: 2025-10-07T13:29:45.677126

Dim oSheet As Sheet
Dim SheetNumber As Double

For Each oSheet In ThisApplication.ActiveDocument.Sheets
SheetNumber  = Mid(oSheet.Name, InStr(1, oSheet.Name, ":") + 1)
MessageBox.Show(Microsoft.VisualBasic.Strings.Format(SheetNumber,"0#"), oSheet.Name)
Next