' Title: Automatic Drawing view sketch, projection, rectangle
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automatic-drawing-view-sketch-projection-rectangle/td-p/13814668
' Category: advanced
' Scraped: 2025-10-09T09:05:50.977766

Sub main()
	
	Dim App As Inventor.Application = ThisApplication
	Dim oDoc As DrawingDocument = App.ActiveDocument
	Dim oTG As TransientGeometry = App.TransientGeometry
	
	Dim oSheet As Sheet = oDoc.Sheets.Item(1) 'only one sheet on doc	
	Dim OnlyView As DrawingView = oSheet.DrawingViews.Item(1) 'only one view on doc. 
	
	Dim ViewSketch As Sketch = OnlyView.Sketches.Add()
	
ViewSketch.Edit() 'You must edit the sketch to work within the view sketch. 
	
	Dim ViewCenterPoint As Point2d = oTG.CreatePoint2d(0, 0) 'The drawing sketch will always place the starting point in the center of the view.
	
	Logger.Info("View width: " & OnlyView.Width & " View height: " & OnlyView.Height)
	
	Dim HalfWidth As Double = OnlyView.Width  / 4
	Dim HalfHeight As Double = OnlyView.Height / 4
	
	Logger.Info("Half Width: " & HalfWidth /2 & " Half height: " & HalfHeight /2)
	
	Dim ViewCornerPoint As SketchPoint = ViewSketch.SketchPoints.Add(oTG.CreatePoint2d(HalfWidth /2,  HalfHeight /2))
	
	ViewSketch.SketchLines.AddAsTwoPointCenteredRectangle(ViewCenterPoint, ViewCornerPoint)
	
ViewSketch.ExitEdit() 'Exit the edit or leave ilogic stuck in sketch. 
	

End Sub