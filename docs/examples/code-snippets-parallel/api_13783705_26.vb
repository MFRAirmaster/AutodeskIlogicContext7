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

Select Case DoorType
	Case "1K" 
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	Case "2K" 
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	Case "3F"
		If DoorWidthOptions <> "3-6"
			Component.IsActive("3F Skin: Back") = True
			Component.IsActive("3F Skin: Front") = True
		Else 
			If DoorHeightOptions = "8-0"
				Component.IsActive("3F (3-6) Skin: Back") = True
				Component.IsActive("3F (3-6) Skin: Front") = True
			Else
				MessageBox.Show(String.Format("DoorType = {0} | DoorWidthOptions = {1} | DoorHeightOptions = {2}", DoorType, DoorWidthOptions, DoorHeightOptions).ToString, "Un-Accounted For Result:")
			End If
		End If
	Case "3S" 
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case "50" 
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	Case Else
		MessageBox.Show(String.Format("DoorType = {0}", DoorType).ToString, "Un-Accounted For DoorType:")
End Select