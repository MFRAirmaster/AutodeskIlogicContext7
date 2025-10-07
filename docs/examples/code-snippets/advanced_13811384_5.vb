' Title: Sample to Modify AssetValue
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sample-to-modify-assetvalue/td-p/13811384#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:35:21.185200

Dim generic_color As Inventor.TextureAssetValue
generic_color = app_ass.Item("hardwood_color")
Dim textureValue As TextureAssetValue = generic_color
Dim texture As AssetTexture = textureValue.Value
Dim textureSubValue As AssetValue = texture.Item("unifiedbitmap_Bitmap")
If texture.Item("texture_WAngle").Value <> 0 Then texture.Item("texture_WAngle").Value = 0