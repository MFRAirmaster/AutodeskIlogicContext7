' Title: Ilogic split a string
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436#messageview_0
' Category: api
' Scraped: 2025-10-07T14:04:35.718094

Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(vbCrLf)

For Each wrd In Words
	'strip out carraige returns & linefeeds
	wrd = wrd.Replace(vbCr, "").Replace(vbLf, "")
	MessageBox.Show("Word Index #" & i & " = " & wrd)
Next