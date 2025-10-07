' Title: Excluding Print and Count if Sheet name matches
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697
' Category: api
' Scraped: 2025-10-07T13:28:29.992552

If oSheet.Name.Contains("Vendor") Then
   oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
   oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
End If