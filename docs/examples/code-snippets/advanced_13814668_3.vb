' Title: Automatic Drawing view sketch, projection, rectangle
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automatic-drawing-view-sketch-projection-rectangle/td-p/13814668#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:03:52.118426

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

Dim oCollection As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection() 

oCollection.Add(oSketch.SketchLines(1)) oCollection.Add(oSketch.SketchLines(2)) oCollection.Add(oSketch.SketchLines(3)) oCollection.Add(oSketch.SketchLines(4)) 

oSketch.OffsetSketchEntitiesUsingDistance(oCollection, 0.1, True, False, true) 

oSketch.DimensionConstraints.AddTwoPointDistance(oSketch.SketchLines(1).EndSketchPoint, oSketch.SketchLines(5).EndSketchPoint, DimensionOrientationEnum.kHorizontalDim, oView.Center)

 oSketch.ExitEdit