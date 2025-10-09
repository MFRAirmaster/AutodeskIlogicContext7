' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-09T08:56:35.499410

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
Select Case DoorType
Case "3F"
	Select Case DoorWidthO
	Case 29.75 To 35.75
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True	
	Case 2 To 29.75
		'Etc Etc Etc
	Case Else
		'Etc Etc Etc
	End Select
Case "3S"
	Select Case DoorWidthO
	Case 12 To 20
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case 2 To 12
		'Etc Etc Etc
	End Select
End Select