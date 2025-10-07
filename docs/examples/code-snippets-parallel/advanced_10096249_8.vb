' Title: Auto ballooning a drawing - Attach balloon to DrawingCurveSegment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-ballooning-a-drawing-attach-balloon-to-drawingcurvesegment/td-p/10096249#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:57:06.094369

Imports Inventor
Module AutoBalloon
#Region "Auto Balloon"
    Private ReadOnly Property _blnViewMargin As Double = 2
    Private ReadOnly Property _blnVerticalOffset As Double = 0.9
    Private ReadOnly Property _blnHorizontalOffset As Double = 1
    Private Property _topLine As LineSegment2d
    Private Property _btmLine As LineSegment2d
    Private Property _rightLine As LineSegment2d
    Private Property _leftLine As LineSegment2d
    Friend Sub AddBallonToView(View As DrawingView)
        Dim PartsList As PartsList = View.Parent.PartsLists.Item(1)
        InitialiseViewBoundingBox(View)
        For Each row As PartsListRow In PartsList.PartsListRows
            If Not row.Ballooned Then CreateRowItemBalloon(row, View)
        Next
        ArrangeBalloonsOnView(View)
    End Sub
    Private Sub InitialiseViewBoundingBox(DView As DrawingView)
        Dim TransGeom As TransientGeometry = InventorApplication.TransientGeometry
        Dim TopLeft As Point2d, TopRight As Point2d, BtmLeft As Point2d, BtmRight As Point2d
        TopLeft = TransGeom.CreatePoint2d(DView.Left, DView.Top)
        TopRight = TransGeom.CreatePoint2d(DView.Left + DView.Width, DView.Top)
        BtmLeft = TransGeom.CreatePoint2d(DView.Left, DView.Top - DView.Height)
        BtmRight = TransGeom.CreatePoint2d(DView.Left + DView.Width, DView.Top - DView.Height)

        _topLine = TransGeom.CreateLineSegment2d(TopLeft, TopRight)
        _btmLine = TransGeom.CreateLineSegment2d(BtmLeft, BtmRight)
        _leftLine = TransGeom.CreateLineSegment2d(TopLeft, BtmLeft)
        _rightLine = TransGeom.CreateLineSegment2d(TopRight, BtmRight)
    End Sub
    Private Sub CreateRowItemBalloon(Item As PartsListRow, View As DrawingView)
        Dim TransGeom As TransientGeometry = InventorApplication.TransientGeometry
        Dim attachPoint As GeometryIntent = GetBalloonAttachGeometry(Item, View)
        If attachPoint Is Nothing Then Exit Sub
        Dim leaderPoint As Point2d = GetBalloonPosition(attachPoint.PointOnSheet, View)
        Dim LeaderPoints As ObjectCollection = InventorApplication.TransientObjects.CreateObjectCollection
        LeaderPoints.Add(leaderPoint)
        LeaderPoints.Add(attachPoint)
        View.Parent.Balloons.Add(LeaderPoints)
    End Sub
    Private Sub ArrangeBalloonsOnView(View As DrawingView)
        Dim leftBalloon As New List(Of Balloon)
        Dim rightBalloon As New List(Of Balloon)
        Dim topBalloon As New List(Of Balloon)
        Dim btmBalloon As New List(Of Balloon)
        For Each balloon As Balloon In View.Parent.Balloons
            If balloon.ParentView Is View Then
                If balloon.Position.X = View.Left - _blnViewMargin Then
                    leftBalloon.Add(balloon)
                ElseIf balloon.Position.X = View.Left + View.Width + _blnViewMargin Then
                    rightBalloon.Add(balloon)
                ElseIf balloon.Position.Y = View.Top + _blnViewMargin Then
                    topBalloon.Add(balloon)
                ElseIf balloon.Position.Y = View.Top - View.Height - _blnViewMargin Then
                    btmBalloon.Add(balloon)
                End If
            End If
        Next
        If leftBalloon.Count > 1 Then ArrangeBalloonVerticaly(leftBalloon)
        If rightBalloon.Count > 1 Then ArrangeBalloonVerticaly(rightBalloon)
        If topBalloon.Count > 1 Then ArrangeBalloonHorizontaly(topBalloon)
        If btmBalloon.Count > 1 Then ArrangeBalloonHorizontaly(btmBalloon)
    End Sub
    Private Sub ArrangeBalloonVerticaly(ballons As List(Of Balloon))
        ballons.Sort(Function(x, y) y.Position.Y.CompareTo(x.Position.Y))
        For index = 1 To ballons.Count - 1
            If ballons(index - 1).Position.Y - ballons(index).Position.Y < _blnVerticalOffset Then
                Dim NewPos As Point2d = ballons(index).Position.Copy
                NewPos.Y = ballons(index - 1).Position.Y - _blnVerticalOffset
                ballons(index).Position = NewPos
            End If
        Next
    End Sub
    Private Sub ArrangeBalloonHorizontaly(ballons As List(Of Balloon))
        ballons.Sort(Function(x, y) y.Position.X.CompareTo(x.Position.X))
        For index = 1 To ballons.Count - 1
            If ballons(index - 1).Position.X - ballons(index).Position.X < _blnHorizontalOffset Then
                Dim NewPos As Point2d = ballons(index).Position.Copy
                NewPos.X = ballons(index - 1).Position.X - _blnHorizontalOffset
                ballons(index).Position = NewPos
            End If
        Next
    End Sub
    Private Function GetBalloonPosition(AttachPoint As Point2d, View As DrawingView) As Point2d

        Dim leaderPoint As Point2d = AttachPoint.Copy
        Dim translationRatio As Double
        Select Case GetQuadrant(AttachPoint, View)
            Case "Top"
                leaderPoint.Y = View.Top + _blnViewMargin
                translationRatio = (leaderPoint.Y - View.Center.Y) / (AttachPoint.Y - View.Center.Y)
                leaderPoint.X = View.Center.X + (AttachPoint.X - View.Center.X) * translationRatio
            Case "Bottom"
                leaderPoint.Y = View.Top - View.Height - _blnViewMargin
                translationRatio = (leaderPoint.Y - View.Center.Y) / (AttachPoint.Y - View.Center.Y)
                leaderPoint.X = View.Center.X + (AttachPoint.X - View.Center.X) * translationRatio
            Case "Left"
                leaderPoint.X = View.Left - _blnViewMargin
                translationRatio = (leaderPoint.X - View.Center.X) / (AttachPoint.X - View.Center.X)
                leaderPoint.Y = View.Center.Y + (AttachPoint.Y - View.Center.Y) * translationRatio
            Case "Right"
                leaderPoint.X = View.Left + View.Width + _blnViewMargin
                translationRatio = (leaderPoint.X - View.Center.X) / (AttachPoint.X - View.Center.X)
                leaderPoint.Y = View.Center.Y + (AttachPoint.Y - View.Center.Y) * translationRatio
            Case "Quadrant not Found"
                Return Nothing
        End Select
        Return leaderPoint
    End Function
    Private Function GetQuadrant(AttachPoint As Point2d, View As DrawingView)
        Dim CornerAngle As Double = Math.Atan2(View.Height, View.Width)
        Dim PointAngle As Double = Math.Atan2(AttachPoint.Y - View.Center.Y, AttachPoint.X - View.Center.X)
        If PointAngle < CornerAngle AndAlso PointAngle > -CornerAngle Then Return "Right"
        If PointAngle > CornerAngle AndAlso PointAngle < Math.PI - CornerAngle Then Return "Top"
        If PointAngle > Math.PI - CornerAngle Or PointAngle < -Math.PI + CornerAngle Then Return "Left"
        If PointAngle > -Math.PI + CornerAngle AndAlso PointAngle < -CornerAngle Then Return "Bottom"
        Return "Quadrant not Found"
    End Function
    Private Function GetBalloonAttachGeometry(Item As PartsListRow, View As DrawingView) As GeometryIntent
        Dim itemOccurrences As ComponentOccurrencesEnumerator = View.ReferencedDocumentDescriptor.ReferencedDocument.ComponentDefinition.Occurrences.AllReferencedOccurrences(Item.ReferencedFiles.Item(1).DocumentDescriptor)
        Dim OccurrencesCurves As List(Of DrawingCurve) = GetCurvesFromOcc(itemOccurrences, View)
        Return GetAttachPoint(GetBestSegmentFromOccurrence(OccurrencesCurves))
    End Function
#End Region
#Region "Curves analysis methods"
    Private Function GetCurvesFromOcc(Occurrences As ComponentOccurrencesEnumerator, View As DrawingView) As List(Of DrawingCurve)
        Dim Curves As New List(Of DrawingCurve)
        For Each occ As ComponentOccurrence In Occurrences
            If Not occ.Suppressed Then
                For Each curve As DrawingCurve In View.DrawingCurves(occ)
                    If Not Curves.Contains(curve) Then Curves.Add(curve)
                Next
            End If
        Next
        Return Curves
    End Function
    Private Function GetBestSegmentFromOccurrence(OccurrenceCurves As List(Of DrawingCurve)) As DrawingCurveSegment
        Dim bestRating As Double = 0, bestSegment As DrawingCurveSegment = Nothing
        For Each curve As DrawingCurve In OccurrenceCurves
            Dim rating As Double
            Dim bestInCurve As DrawingCurveSegment = GetBestSegmentFromCurve(curve, rating)
            If rating > bestRating Then bestSegment = bestInCurve : bestRating = rating
        Next
        Return bestSegment
    End Function
    Private Function GetBestSegmentFromCurve(Curve As DrawingCurve, ByRef segmentRating As Double) As DrawingCurveSegment
        Dim bestSegment As DrawingCurveSegment = Nothing
        segmentRating = 0
        For Each segment As DrawingCurveSegment In Curve.Segments
            Dim segLength As Double = GetSegmentLength(segment)
            Dim segmentPoints As List(Of Point2d) = SplitSegment(segment, 10)
            Dim closestDist As Double = Double.PositiveInfinity
            Dim distanceSum As Double = 0
            For Each point As Point2d In segmentPoints
                Dim pointDist As Double = DistanceToViewRangeBox(point, Curve.Parent)
                If pointDist < closestDist Then closestDist = pointDist
                distanceSum += pointDist
            Next
            Dim avgDistance = distanceSum / segmentPoints.Count
            Dim rating As Double = segLength ' / (avgDistance * closestDist ^ 2)
            If Curve.EdgeType = DrawingEdgeTypeEnum.kTangentEdge Then rating /= 10
            If rating > segmentRating Then bestSegment = segment : segmentRating = rating
        Next
        Return bestSegment
    End Function
    Private Function SplitSegment(Segment As DrawingCurveSegment, SplitPrecision As Integer) As List(Of Point2d)
        Dim pointList As New List(Of Point2d), MinParam, MaxParam As Double, TransGeom As TransientGeometry
        TransGeom = InventorApplication.TransientGeometry
        Segment.Geometry.Evaluator.GetParamExtents(MinParam, MaxParam)
        For i = 0 To SplitPrecision
            Dim pCoordinate(0 To 1) As Double
            Segment.Geometry.Evaluator.GetPointAtParam({MinParam + i * (MaxParam - MinParam) / SplitPrecision}, pCoordinate)
            pointList.Add(TransGeom.CreatePoint2d(pCoordinate(0), pCoordinate(1)))
        Next
        Return pointList
    End Function
    Private Function DistanceToViewRangeBox(TestPoint As Point2d, DView As DrawingView) As Double

        Dim shortest As Double = _topLine.DistanceTo(TestPoint)
        Dim dist As Double = _btmLine.DistanceTo(TestPoint)
        If dist < shortest Then shortest = dist
        dist = _leftLine.DistanceTo(TestPoint)
        If dist < shortest Then shortest = dist
        dist = _leftLine.DistanceTo(TestPoint)
        If dist < shortest Then shortest = dist
        Return shortest
    End Function
    Private Function GetSegmentLength(Segment As DrawingCurveSegment) As Double
        Select Case Segment.GeometryType
            Case Curve2dTypeEnum.kLineCurve2d
                Return Segment.EndPoint.DistanceTo(Segment.StartPoint)
            Case Else
                Return GetCurveLength(Segment.Geometry.Evaluator)
        End Select
    End Function
    Private Function GetSegmentMidPoint(Segment As DrawingCurveSegment) As Point2d
        Dim CurveEval As Curve2dEvaluator = Segment.Geometry.Evaluator
        Dim minParam, maxParam, midParam(0), midPointCoordinates(1), curveLength As Double
        CurveEval.GetParamExtents(minParam, maxParam)
        CurveEval.GetLengthAtParam(minParam, maxParam, curveLength)
        CurveEval.GetParamAtLength(minParam, curveLength / 2, midParam(0))
        CurveEval.GetPointAtParam(midParam, midPointCoordinates)
        Return InventorApplication.TransientGeometry.CreatePoint2d(midPointCoordinates(0), midPointCoordinates(1))
    End Function
    Private Function GetCurveLength(Eval As Curve2dEvaluator) As Double
        Dim minParam As Double, maxParam As Double, curveLength As Double
        Eval.GetParamExtents(minParam, maxParam)
        Eval.GetLengthAtParam(minParam, maxParam, curveLength)
        Return curveLength
    End Function
    Private Function GetAttachPoint(Segment As DrawingCurveSegment) As GeometryIntent
        If Segment Is Nothing Then Return Nothing
        Dim DrawingSheet As Sheet = Segment.Parent.Parent.Parent
        If Segment.GeometryType = Curve2dTypeEnum.kCircleCurve2d Or Segment.GeometryType = Curve2dTypeEnum.kEllipseFullCurve2d Then
            Select Case GetQuadrant(Segment.Geometry.Center, Segment.Parent.Parent)
                Case "Top"
                    Return DrawingSheet.CreateGeometryIntent(Segment.Parent, PointIntentEnum.kCircularTopPointIntent)
                Case "Bottom"
                    Return DrawingSheet.CreateGeometryIntent(Segment.Parent, PointIntentEnum.kCircularBottomPointIntent)
                Case "Left"
                    Return DrawingSheet.CreateGeometryIntent(Segment.Parent, PointIntentEnum.kCircularLeftPointIntent)
                Case "Right"
                    Return DrawingSheet.CreateGeometryIntent(Segment.Parent, PointIntentEnum.kCircularRightPointIntent)
                Case "Quadrant not Found"
                    Return DrawingSheet.CreateGeometryIntent(Segment.Parent, PointIntentEnum.kCenterPointIntent)
            End Select
        End If
        Return DrawingSheet.CreateGeometryIntent(Segment.Parent, GetSegmentMidPoint(Segment))
    End Function
#End Region
End Module