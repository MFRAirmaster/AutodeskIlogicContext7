' Title: Section &amp; Detail View Label
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080#messageview_0
' Category: api
' Scraped: 2025-10-09T08:53:40.602390

Public Function GetAlphaString(ByVal ColumnLetter As String) As String
	'convert to number
	For i = 1 To Len(ColumnLetter)
		AlphaNum = AlphaNum * 26 + (Asc(UCase(Mid(ColumnLetter, i, 1))) -64)
	Next

	AlphaNum = AlphaNum + 1

	'convert back to format
	Do While AlphaNum > 0
		AlphaNum = AlphaNum - 1
		GetAlphaString = ChrW(65 + AlphaNum Mod 26) & GetAlphaString
		AlphaNum = AlphaNum \ 26
	Loop
End Function