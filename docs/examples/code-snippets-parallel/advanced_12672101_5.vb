' Title: Get overall dimensions  of all views of a drawing by iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/get-overall-dimensions-of-all-views-of-a-drawing-by-ilogic/td-p/12672101
' Category: advanced
' Scraped: 2025-10-07T14:35:53.111302

Public Class ThisRule
    Private _doc As DrawingDocument
    Private _sheet As Sheet
    Private _view As DrawingView
    Private _intents As List(Of GeometryIntent) = New List(Of GeometryIntent)()

    Sub Main()
        _doc = ThisDoc.Document
		
		For Each sheet As Sheet In _doc.Sheets
            For Each view As DrawingView In Sheet.DrawingViews
                If (View.ViewType <> DrawingViewTypeEnum.kStandardDrawingViewType) Then Continue For
				_intents = New List(Of GeometryIntent)()
				_sheet = Sheet
				_view = View
				CreateIntentList()
		        createHorizontalOuterDimension()
		        createVerticalOuterDimension()
            Next
        Next
    End Sub

    Private Sub createHorizontalOuterDimension()
        Dim orderedIntents = _intents.OrderByDescending(Function(s) s.PointOnSheet.X)

        Dim pointLeft = orderedIntents.First
        Dim pointRight = orderedIntents.Last

        Dim textX = pointLeft.PointOnSheet.X +
                (pointRight.PointOnSheet.X - pointLeft.PointOnSheet.X) / 2
        Dim textY = _view.Position.Y + _view.Height / 2 + 2

        Dim pointText = ThisApplication.TransientGeometry.CreatePoint2d(textX, textY)
        _sheet.DrawingDimensions.GeneralDimensions.AddLinear(
            pointText, pointLeft, pointRight, DimensionTypeEnum.kHorizontalDimensionType)
    End Sub
    Private Sub createVerticalOuterDimension()
        Dim orderedIntents = _intents.OrderByDescending(Function(s) s.PointOnSheet.Y)

        Dim pointLeft = orderedIntents.Last
        Dim pointRight = orderedIntents.First

        Dim textY = pointLeft.PointOnSheet.Y +
                (pointRight.PointOnSheet.Y - pointLeft.PointOnSheet.Y) / 2
        Dim textX = _view.Position.X - _view.Width / 2 - 2

        Dim pointText = ThisApplication.TransientGeometry.CreatePoint2d(textX, textY)
        _sheet.DrawingDimensions.GeneralDimensions.AddLinear(
            pointText, pointLeft, pointRight, DimensionTypeEnum.kVerticalDimensionType)
    End Sub

    Private Sub addIntent(Geometry As DrawingCurve, IntentPlace As Object, onLineCheck As Boolean)
        Dim intent As GeometryIntent = _sheet.CreateGeometryIntent(Geometry, IntentPlace)
        If intent.PointOnSheet Is Nothing Then Return

        If onLineCheck Then
            If (IntentIsOnCurve(intent)) Then
                _intents.Add(intent)
            End If
        Else
            _intents.Add(intent)
        End If
    End Sub

    Private Function IntentIsOnCurve(intent As GeometryIntent) As Boolean
        Dim Geometry As DrawingCurve = intent.Geometry
        Dim sp = intent.PointOnSheet

        Dim pts(1) As Double
        Dim gp() As Double = {}
        Dim md() As Double = {}
        Dim pm() As Double = {}
        Dim st() As SolutionNatureEnum = {}
        pts(0) = sp.X
        pts(1) = sp.Y

        Try
            Geometry.Evaluator2D.GetParamAtPoint(pts, gp, md, pm, st)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Sub CreateIntentList()

        For Each oDrawingCurve As DrawingCurve In _view.DrawingCurves
            Select Case oDrawingCurve.ProjectedCurveType
                Case _
                        Curve2dTypeEnum.kCircleCurve2d,
                        Curve2dTypeEnum.kCircularArcCurve2d,
                        Curve2dTypeEnum.kEllipseFullCurve2d,
                        Curve2dTypeEnum.kEllipticalArcCurve2d

                    addIntent(oDrawingCurve, PointIntentEnum.kCircularTopPointIntent, True)
                    addIntent(oDrawingCurve, PointIntentEnum.kCircularBottomPointIntent, True)
                    addIntent(oDrawingCurve, PointIntentEnum.kCircularLeftPointIntent, True)
                    addIntent(oDrawingCurve, PointIntentEnum.kCircularRightPointIntent, True)

                    addIntent(oDrawingCurve, PointIntentEnum.kEndPointIntent, False)
                    addIntent(oDrawingCurve, PointIntentEnum.kStartPointIntent, False)

                Case _
                        Curve2dTypeEnum.kLineCurve2d,
                        Curve2dTypeEnum.kLineSegmentCurve2d

                    addIntent(oDrawingCurve, PointIntentEnum.kEndPointIntent, False)
                    addIntent(oDrawingCurve, PointIntentEnum.kStartPointIntent, False)

                Case _
                    Curve2dTypeEnum.kPolylineCurve2d,
                    Curve2dTypeEnum.kBSplineCurve2d,
                    Curve2dTypeEnum.kUnknownCurve2d

                    ' Unhandled curves types
                Case Else
            End Select
        Next
    End Sub

    ' Copyright 2021
    ' 
    ' This code was written by Jelte de Jong, and published on www.hjalte.nl
    '
    ' Permission Is hereby granted, free of charge, to any person obtaining a copy of this 
    ' software And associated documentation files (the "Software"), to deal in the Software 
    ' without restriction, including without limitation the rights to use, copy, modify, merge, 
    ' publish, distribute, sublicense, And/Or sell copies of the Software, And to permit persons 
    ' to whom the Software Is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice And this permission notice shall be included In all copies Or
    ' substantial portions Of the Software.
    ' 
    ' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or IMPLIED, 
    ' INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY, FITNESS For A PARTICULAR 
    ' PURPOSE And NONINFRINGEMENT. In NO Event SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE 
    ' For ANY CLAIM, DAMAGES Or OTHER LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or 
    ' OTHERWISE, ARISING FROM, OUT Of Or In CONNECTION With THE SOFTWARE Or THE USE Or OTHER 
    ' DEALINGS In THE SOFTWARE.
End Class