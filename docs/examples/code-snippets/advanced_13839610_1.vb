' Title: 3D sketch line between two UCS center points
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/3d-sketch-line-between-two-ucs-center-points/td-p/13839610#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:29:37.139546

Sub main()
	
	Dim App As Inventor.Application = ThisApplication
	Dim oDoc As PartDocument = App.ActiveDocument
	Dim oCompDef As PartComponentDefinition = oDoc.ComponentDefinition
	
	Dim oTG As TransientGeometry = App.TransientGeometry
	
	Dim UCS1 As UserCoordinateSystem
	Dim UCS2 As UserCoordinateSystem
	
	For Each iUCS As UserCoordinateSystem In oCompDef.UserCoordinateSystems 'Collect the UCS objects dynamically, This example uses their names. 
		If iUCS.Name.Contains("First UCS") Then 
			UCS1 = iUCS
		Else If iUCS.Name.Contains("Second UCS") Then 
			UCS2 = iUCS
		End If 
	Next
		
	Dim UCSPoint1 As Point = UCS1.Origin.Point 'Throw the points of the origins into usable point objects. 
	Dim UCSPoint2 As Point = UCS2.Origin.Point
	
	Dim UCSPoint1B As WorkPoint = oTG.CreatePoint(UCS1.Origin.Point.X, UCS1.Origin.Point.Y, UCS1.Origin.Point.Z) 'How to create different kinds of points depending on needs. 
	Dim UCSPoint2B As WorkPoint = oTG.CreatePoint(UCS2.Origin.Point.X, UCS2.Origin.Point.Y, UCS2.Origin.Point.Z)
	
	
	Dim Sketch3d As Sketch3D = oCompDef.Sketches3D.Add()


	Dim SketchLine As SketchLine3D
	
	Try
		SketchLine	= Sketch3d.SketchLines3D.AddByTwoPoints(UCSPoint1, UCSPoint2)
	Catch
		SketchLine = Sketch3d.SketchLines3D.AddByTwoPoints(UCSPoint1B, UCSPoint2B) 'The .addbytwo points calls for general point objects. Sometimes finniky, you might need to supply different points. 
	End Try
	
	
	
	
End Sub