' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436
' Category: api
' Scraped: 2025-10-07T12:41:55.729035

Dim Separators() As Char = {";"} 
Sentence = "1;2;3;4;5"
Words = Sentence.Split(Separators)
i = 0
For Each wrd In Words
MessageBox.Show("Word Index #" & i & " = " & Words(i))
i=i+1
Next