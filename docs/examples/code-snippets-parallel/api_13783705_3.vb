' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

If DoorType = "3F" And DoorWidthOptions = ("2-6" or "2-8" or "2-10" or "3-0") Then
		Component.IsActive("1K Skin: Back") = False
		Component.IsActive("1K Skin: Front") = False
		Component.IsActive("2K Skin: Back") = False
		Component.IsActive("2K Skin: Front") = False
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
		Component.IsActive("3F (3-6) Skin: Back") = False
		Component.IsActive("3F (3-6) Skin: Front") = False
		Component.IsActive("3S Skin: Back") = False
		Component.IsActive("3S Skin: Front") = False
		Component.IsActive("50 Skin: Back") = False
		Component.IsActive("50 Skin: Front") = False
	End If