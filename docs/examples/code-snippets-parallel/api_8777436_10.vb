' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436#messageview_0
' Category: api
' Scraped: 2025-10-07T14:04:35.718094

Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Separators() As Char = {".", "-", " " }
Dim Words() As String = Sentence.Split(Separators)
Dim Word1 As String = Words(0)
Dim Word2 As String = Words(1)
Dim Word3 As String = Words(2)
Dim Word4 As String = Words(3)
Dim Word5 As String = Words(4)
Dim Word6 As String = Words(5)
MsgBox("Word1 = " & Word1 & _
vbCrLf & "Word2 = " & Word2 & _
vbCrLf & "Word3 = " & Word3 & _
vbCrLf & "Word4 = " & Word4 & _
vbCrLf & "Word5 = " & Word5 & _
vbCrLf & "Word6 = " & Word6, vbInformation, "Sentence Split Into Words")