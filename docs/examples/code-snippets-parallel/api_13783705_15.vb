' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

If DoorType = "3F" Then
	If DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
ElseIf DoorType = "X" Then
	'Etc etc etc
ElseIf DoorType = "Y" Then
	'Etc etc etc
ElseIf DoorType = "Z" Then
	'Etc etc etc
End If