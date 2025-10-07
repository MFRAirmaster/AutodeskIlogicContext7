' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436#messageview_0
' Category: api
' Scraped: 2025-10-07T14:04:35.718094

Dim Separators() As Char = {";"} 
Sentence = "10;20;3;4;5"
Words = Sentence.Split(Separators)

spot1 = Words(0)
spot2 = Words(1)

Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True