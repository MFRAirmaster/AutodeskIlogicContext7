' Title: How can I annotate these dimensions in a drawing, using iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-can-i-annotate-these-dimensions-in-a-drawing-using-ilogic/td-p/13738809
' Category: advanced
' Scraped: 2025-10-07T13:11:30.425074

Public Sub Main()
  Dim oDrawDoc As DrawingDocument
  oDrawDoc = ThisApplication.ActiveDocument

  Dim oActiveSheet As Sheet
  oActiveSheet = oDrawDoc.ActiveSheet

  Dim oDrawingCurveSegment As DrawingCurveSegment
  oDrawingCurveSegment = ThisApplication.CommandManager.Pick(kDrawingCurveSegmentFilter, "Seleccionar la línea del origen")

  Dim oDrawingCurve As DrawingCurve
  oDrawingCurve = oDrawingCurveSegment.Parent

  If Not oDrawingCurve.CurveType = kLineSegmentCurve Then
    MsgBox("Esta no es una línea Ingeniero")
    Exit Sub
  End If

  Dim oDimIntent As GeometryIntent
  oDimIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kEndPointIntent)

  Dim oDrawingView As DrawingView
  oDrawingView = oDrawingCurve.Parent

  If Not oDrawingView.HasOriginIndicator Then
    oDrawingView.CreateOriginIndicator(oDimIntent)
  End If

  Dim oOrdinateDimensions As OrdinateDimensions
  oOrdinateDimensions = oActiveSheet.DrawingDimensions.OrdinateDimensions

  Dim oTextOrigin As Point2d
  Dim DimType As DimensionTypeEnum
  DimType = kHorizontalDimensionType

  Dim TG As TransientGeometry
  TG = ThisApplication.TransientGeometry

  oTextOrigin = TG.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.StartPoint.Y)
  Call oOrdinateDimensions.Add(oDimIntent, oTextOrigin, DimType)

  Dim oIntent As GeometryIntent
  Dim textPt As Point2d

  For Each oDrawingCurve In oDrawingView.DrawingCurves
    Select Case oDrawingCurve.CurveType
      Case kCircleCurve
        oIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kCenterPointIntent)
        textPt = TG.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.CenterPoint.Y)

      Case kLineSegmentCurve
        oIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kStartPointIntent)
        textPt = TG.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.StartPoint.Y)

      Case kCircularArcCurve
        oIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kEndPointIntent)
        textPt = TG.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.EndPoint.Y)

      Case Else
        oIntent = Nothing
    End Select

    If Not oIntent Is Nothing Then
      Call oOrdinateDimensions.Add(oIntent, textPt, DimType)
    End If
  Next

  oDrawingCurveSegment = ThisApplication.CommandManager.Pick(kDrawingCurveSegmentFilter, "Select end line to complete ordinate dimension")
  oDrawingCurve = oDrawingCurveSegment.Parent

  If Not oDrawingCurve.CurveType = kLineSegmentCurve Then
    MsgBox("A linear curve should be selected for this sample.")
    Exit Sub
  End If

  oDimIntent = oActiveSheet.CreateGeometryIntent(oDrawingCurve, kEndPointIntent)
  oTextOrigin = TG.CreatePoint2d(oDrawingView.Left + 2, oDrawingCurve.StartPoint.Y)
  Call oOrdinateDimensions.Add(oDimIntent, oTextOrigin, DimType)

End Sub