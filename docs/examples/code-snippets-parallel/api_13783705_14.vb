' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

If DoorType = "3F" And DoorWidth <= 40 Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth >= 40  And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If