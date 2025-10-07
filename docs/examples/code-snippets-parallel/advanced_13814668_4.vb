' Title: Automatic Drawing view sketch, projection, rectangle
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automatic-drawing-view-sketch-projection-rectangle/td-p/13814668#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:25:56.965551

Dim InventorApp As Inventor.Application = ThisApplication
Dim oDrawingDoc As DrawingDocument = InventorApp.ActiveDocument

Dim activeSheet As Sheet
Dim oView As DrawingView

oSheet = oDrawingDoc.ActiveSheet
oView = oSheet.DrawingViews(1)

Dim oSketch As DrawingSketch
Dim oCurve As DrawingCurve


oSketch = oView.Sketches.Add()
oSketch.Edit

For Each oCurve In oView.DrawingCurves
	Try
		oSketch.AddByProjectingEntity(oCurve)
	Catch
	End Try
Next

offset = 0.1 in
x = (oView.Width / 2 / oView.Scale) + offset
y = oView.Height / 2 / oView.Scale + offset
pt1 = ThisApplication.TransientGeometry.CreatePoint2d(-x, -y)
pt2 = ThisApplication.TransientGeometry.CreatePoint2d(x, y)
textPt = ThisApplication.TransientGeometry.CreatePoint2d(x+1, y+1)

oSketch.SketchLines.AddAsTwoPointRectangle(pt1, pt2)

Dim Dimension1 = oSketch.DimensionConstraints.AddTwoPointDistance(oSketch.SketchLines(1).EndSketchPoint,
oSketch.SketchLines(5).EndSketchPoint, DimensionOrientationEnum.kHorizontalDim, textPt)
oSketch.ExitEdit