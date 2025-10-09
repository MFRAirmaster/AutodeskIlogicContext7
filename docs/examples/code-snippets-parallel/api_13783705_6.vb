' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

'"3F Skin" ActiveIf DoorType = "3F" And DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If
'"3F (3-6) Skin" Active
If DoorType = "3F" And DoorWidthOptions = "3-6" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If