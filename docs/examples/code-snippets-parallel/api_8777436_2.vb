' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436
' Category: api
' Scraped: 2025-10-09T09:09:57.471675

Dim Separators() As Char = {";"} 
Sentence = "10;20;3;4;5"
Words = Sentence.Split(Separators)

spot1 = Words(0)
spot2 = Words(1)

Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True