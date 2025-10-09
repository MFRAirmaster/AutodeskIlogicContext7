' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	End If

If DoorType = "2K" Then
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "3-0" Then
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

If DoorType = "3F" And DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If

If DoorType = "3S" Then
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	End If

If DoorType = "50" Then
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	End If