' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T13:18:13.978278

Dim A,B,C As String
A = "1"
B = "0"
C = "0"

If A = "1" And (B = "1" Or C = "1")
	MessageBox.Show("1", "Test")
Else
	MessageBox.Show("2", "Test")
end if