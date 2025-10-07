' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436
' Category: api
' Scraped: 2025-10-07T12:41:55.729035

Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Words() As String = Sentence.Split({".", "-", " " },StringSplitOptions.RemoveEmptyEntries)
Dim sResults As String
For i As Integer = 0 To Words.Length - 1
	sResults &= "Word " & i & " = " & Words(i) & vbCrLf
Next
MsgBox(sResults, vbInformation, "Sentence Split Into Words")