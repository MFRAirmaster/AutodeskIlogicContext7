' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T13:18:13.978278

If DoorType = "3F" And DoorWidthOptions = "3-0" or "2-10" or "2-8" or "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End IfIf DoorType = "3F" And DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then                Component.IsActive("3F Skin: Back") = True                Component.IsActive("3F Skin: Front") = True        End If