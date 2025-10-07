' Title: Can I rotate radius general dimension in drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/can-i-rotate-radius-general-dimension-in-drawing/td-p/13757148#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:32:41.363431

Sub zzz()

Dim doc As DrawingDocument
Set doc = ThisApplication.ActiveDocument

Dim oSheet As sheet
Set oSheet = doc.activeSheet

Dim oDim As RadiusGeneralDimension
Set oDim = ThisApplication.CommandManager.Pick(kDrawingDimensionFilter, "Pick dimension")

Dim newDim As RadiusGeneralDimension

Dim oTG As TransientGeometry
Set oTG = ThisApplication.TransientGeometry

Dim newGeometryIntent As GeometryIntent
Set newGeometryIntent = oDim.Intent '<--- stuck here

End sub