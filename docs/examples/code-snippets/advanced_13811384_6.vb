' Title: Sample to Modify AssetValue
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sample-to-modify-assetvalue/td-p/13811384#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:35:21.185200

Dim material As MaterialAsset
material = projekt_asslib.MaterialAssets("mat name")
material.PhysicalPropertiesAsset.Item("structural_Density").Value=2
Logger.Info(material.PhysicalPropertiesAsset.Item("structural_Density").Value)