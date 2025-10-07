' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436#messageview_0
' Category: api
' Scraped: 2025-10-07T14:04:35.718094

Dim Separators() As Char = {";"} 
Sentence = StringParameter
If Sentence = "1;1" Then
Words = Sentence.Split(Separators)
spot1 = Words(0)
spot2 = Words(1)
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
Else If Sentence = "1;5;7" Then
Words = Sentence.Split(Separators)
spot1 = Words(0)
spot2 = Words(1)
spot3 = Words(2)
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
End If 
InventorVb.DocumentUpdate()