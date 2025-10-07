' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T14:07:55.595605

Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F"
	Select Case DoorWidthOptions
	Case 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 22
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 20
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 18
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select