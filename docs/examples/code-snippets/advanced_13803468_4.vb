' Title: SurfaceBody Appearance Override ignored?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/surfacebody-appearance-override-ignored/td-p/13803468
' Category: advanced
' Scraped: 2025-10-07T12:33:41.827885

' Assign a color to the node using an existing appearance asset.
Dim oAppearance As Asset = oDoc.AppearanceAssets(1)
oLineNode.Appearance = oAppearance