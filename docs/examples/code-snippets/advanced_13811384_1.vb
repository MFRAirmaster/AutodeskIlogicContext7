' Title: Sample to Modify AssetValue
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sample-to-modify-assetvalue/td-p/13811384#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:35:21.185200

Dim doc As Document
    doc = ThisApplication.ActiveDocument
    
    ' Only document appearances can be edited, so that's what's created.
    ' This assumes a part or assembly document is active.
    Dim docAssets As Assets
    docAssets = doc.Assets
    
    ' Create a new appearance asset.
    Dim appearance As Asset
    appearance = docAssets.Add(kAssetTypeAppearance, "Generic", _
                                    "MyShinyRed", "My Shiny Red Color")
    
    Dim tobjs As TransientObjects
    tobjs = ThisApplication.TransientObjects


    Dim color As ColorAssetValue
    color = appearance.Item("generic_diffuse")
    color.value = tobjs.CreateColor(255, 15, 15)
    
    Dim floatValue As FloatAssetValue
    floatValue = appearance.Item("generic_reflectivity_at_0deg")
    floatValue.value = 0.5
    
    floatValue = appearance.Item("generic_reflectivity_at_90deg")
    floatValue.value = 0.5