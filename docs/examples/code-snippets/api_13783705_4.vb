' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T13:18:13.978278

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
If DoorType = "3F" And DoorWidthO >= 29.75 And DoorWidthO <= 35.75
	Component.IsActive("3F Skin: Back") = True
	Component.IsActive("3F Skin: Front") = True
ElseIf DoorType = "3S" And DoorWidthO >= 29.75 And DoorWidthO <= 35.75
	Component.IsActive("3S Skin: Back") = True
	Component.IsActive("3S Skin: Front") = True
End If