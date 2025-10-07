' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T13:18:13.978278

Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F", "3F Skin", "3F (3-6) Skin"
	Select Case DoorWidthOptions
	Case 18, 20, 22, 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select