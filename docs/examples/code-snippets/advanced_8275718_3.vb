' Title: Auto Ordinate Dimensions
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-ordinate-dimensions/td-p/8275718#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:45:52.795768

Public Sub Main()
  ' Set a reference to the drawing document.
  ' This assumes a drawing document is active.
  Dim oDrawDoc As DrawingDocument
    oDrawDoc = ThisApplication.ActiveDocument

  ' Set a reference to the active sheet.
  Dim oActiveSheet As Sheet
    oActiveSheet = oDrawDoc.ActiveSheet

  '   a reference to the drawing curve segment.
  ' This assumes that a linear drawing curve is selected.
  Dim oDrawingCurveSegment As DrawingCurveSegment
    oDrawingCurveSegment = ThisApplication.CommandManager.Pick(kDrawingCurveSegmentFilter, "Select start line to start ordinate dimension")

  '   a reference to the drawing curve.
  Dim oDrawingCurve As DrawingCurve
    oDrawingCurve = oDrawingCurveSegment.Parent

  If Not oDrawingCurve.CurveType = kLineSegmentCurve Then
    MsgBox ("A linear curve should be selected for this sample.")
    Exit Sub
  End If

  ' Create point intents to anchor the dimension to.
  Dim oDimIntent  As GeometryIntent
    oDimIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kEndPointIntent)

  '   a reference to the view to which the curve belongs.
  Dim oDrawingView As DrawingView
    oDrawingView = oDrawingCurve.Parent

  ' If origin indicator has not been already created, create it first.
  If Not oDrawingView.HasOriginIndicator Then
    ' The indicator will be located at the start point of the selected curve.
    oDrawingView.CreateOriginIndicator (oDimIntent)
  End If

  '   a reference to the ordinate dimensions collection.
  Dim oOrdinateDimensions As OrdinateDimensions
    oOrdinateDimensions = oActiveSheet.DrawingDimensions.OrdinateDimensions

  Dim oTextOrigin  As Point2d
  Dim DimType As DimensionTypeEnum
  
  ' Selected curve is vertical or at an angle.
  DimType = kHorizontalDimensionType
    
  '   the text points for the 2 dimensions.
    oTextOrigin = ThisApplication.TransientGeometry.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.StartPoint.Y)
 
  ' Create the first ordinate dimension.
  Call oOrdinateDimensions.Add(oDimIntent, oTextOrigin, DimType)
  
  For Each oDrawingCurve In oDrawingView.DrawingCurves
    If oDrawingCurve.CurveType = kCircleCurve Then
        Dim oIntent As GeometryIntent
          oIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kCenterPointIntent)
        
        Dim origin As Point2d
          origin = ThisApplication.TransientGeometry.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.CenterPoint.Y)
        
        Call oOrdinateDimensions.Add(oIntent, origin, DimType)
    End If
  Next
  
    oDrawingCurveSegment = ThisApplication.CommandManager.Pick(kDrawingCurveSegmentFilter, "Select end line to complete ordinate dimension")
  
    oDrawingCurve = oDrawingCurveSegment.Parent
  
  If Not oDrawingCurve.CurveType = kLineSegmentCurve Then
    MsgBox ("A linear curve should be selected for this sample.")
    Exit Sub
  End If
  
    oDimIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kEndPointIntent)
  
    oTextOrigin = ThisApplication.TransientGeometry.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.StartPoint.Y)
  
  Call oOrdinateDimensions.Add(oDimIntent, oTextOrigin, DimType)
  
End Sub