' Title: Retrieve 3D dimensions placed on an assembly in a drawing without sketches dimensions.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/retrieve-3d-dimensions-placed-on-an-assembly-in-a-drawing/td-p/10901052#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:39:30.326321

oDDoc = ThisDrawing.Document
oSheet = oDDoc.ActiveSheet
oView = oSheet.DrawingViews.Item(1)
oView.Include3DAnnotations = True