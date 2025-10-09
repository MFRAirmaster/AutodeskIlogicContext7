' Title: create a centerline by two points
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-a-centerline-by-two-points/td-p/12710357#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:06:29.514764

Dim drawView As DrawingView = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a drawing view")

Dim drawDoc As DrawingDocument= ThisApplication.ActiveDocument

Dim sketch1 As DrawingSketch = drawView.Sketches.Add   
sketch1.Edit

Dim startPt As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(0,0)
startPt = drawView.SheetToDrawingViewSpace(startPt)

Dim endPt As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(216*2.54,0)
endPt = drawView.SheetToDrawingViewSpace(endPt)

Dim oLine As SketchLine = sketch1.SketchLines.AddByTwoPoints(startPt, endPt)
Dim layers As LayersEnumerator = drawDoc.StylesManager.Layers
oLine.Layer = layers.Item("Centerline (ANSI)")
sketch1.ExitEdit