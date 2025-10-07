' Title: Auto Ordinate Dimensions
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-ordinate-dimensions/td-p/8275718#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:28:13.380673

Sub Main()
	
	Dim oDrawDoc As DrawingDocument
	oDrawDoc = ThisApplication.ActiveDocument
	
	Dim oActiveSheet As Sheet
	oActiveSheet = oDrawDoc.ActiveSheet
	
	Dim oTG As TransientGeometry
	oTG = ThisApplication.TransientGeometry
	
	Dim oDimIntent As GeometryIntent

	Dim Entity As DrawingCurveSegment
	Entity = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select a kDrawingCurveSegmentFilter")
	
	' Get all the selected drawing curve segments.
	Dim oDrawingCurve As DrawingCurve
	oDrawingCurve = Entity.Parent
	
	Dim oDrawingView As DrawingView
	oDrawingView = oDrawingCurve.Parent
	
	Dim oDCE As DrawingCurvesEnumerator
	oDCE = oDrawingView.DrawingCurves()
	
	Dim oODS As OrdinateDimensions
	oODS = oActiveSheet.DrawingDimensions.OrdinateDimensions
	
	Dim oOD As OrdinateDimension
	
	Dim j As Integer
	Dim CurveTypeCount(3) As Integer
	CurveTypeCount(1)=0
	CurveTypeCount(2)=0
	CurveTypeCount(3)=0
		
	For j = 1 To oDCE.Count()
	
		oDrawingCurve = oDCE.Item(j)
		
		Select Case oDrawingCurve.ProjectedCurveType()
			Case 5252 'kCircleCurve2d 
				CurveTypeCount(1)=CurveTypeCount(1) + 1
			Case 5253 'kCircularArcCurve2d 
				CurveTypeCount(2)=CurveTypeCount(2) + 1
			Case 5251 'kLineSegmentCurve2d 
				CurveTypeCount(3)=CurveTypeCount(3) + 1
		End Select
		
		If oDrawingCurve.ProjectedCurveType() = 5252  Then'kCircleCurve2d
			
			MessageBox.Show("oDrawingCurve.Type())" & oDrawingCurve.Type(), "kCircleCurve2d") ' kDrawingCurveObject 
			
			oDimIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, oDrawingCurve)

			Dim oTextOrigin1 As Point2d
			oTextOrigin1 = oTG.CreatePoint2d(5, 5)
			MessageBox.Show("oTextOrigin1 initialized")
			
			oOD = oODS.Add( kCenterPointIntent, oTextOrigin1, kHorizontalDimensionType)
			MessageBox.Show("HURRAY IT WORKED!!!!!!!!!!!!")
		End If
	Next
End Sub