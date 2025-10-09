' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F"
	Select Case DoorWidthOptions
	Case 18 to 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select