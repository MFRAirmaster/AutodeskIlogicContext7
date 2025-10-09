' Title: Auto ballooning a drawing - Attach balloon to DrawingCurveSegment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-ballooning-a-drawing-attach-balloon-to-drawingcurvesegment/td-p/10096249
' Category: advanced
' Scraped: 2025-10-09T09:00:08.332257

Private Oapp As Inventor.Application
    Private TransGeom As TransientGeometry

    Public Sub New(inventorApp As Inventor.Application)
        Oapp = inventorApp
        TransGeom = Oapp.TransientGeometry
    End Sub

    Public Sub Main(oDrawingDoc As DrawingDocument, oSheet As Sheet, oView As DrawingView, sectionview As DrawingView)
        Dim sw As New Stopwatch()
        sw.Start()

        If oView IsNot Nothing Then
            Dim sheet2 As Sheet = oDrawingDoc.Sheets.Item(2)
            Dim partlist As PartsList = sheet2.PartsLists.Item(1)

            AddBalloonToView(oView, partlist)
            AddBalloonToView(sectionview, partlist)
        End If

        sw.Stop()
        Dim elapsedMinutes As Integer = sw.Elapsed.Minutes
        Dim elapsedSeconds As Integer = sw.Elapsed.Seconds
        Dim message As String = $"Balloon generation completed in {elapsedMinutes} minute(s) and {elapsedSeconds} second(s)."
        System.Windows.Forms.MessageBox.Show(message, "Balloon Generation Time")

    End Sub

    Private Sub AddBalloonToView(View As DrawingView, partlist1 As PartsList)
        InitialiseViewBoundingBox(View)

        For Each row As PartsListRow In partlist1.PartsListRows
            If Not row.Ballooned AndAlso row.Visible Then
                CreateRowItemBalloon(row, View)
            End If
        Next

        ArrangeBalloonsAroundView(View)
    End Sub

    Private Sub InitialiseViewBoundingBox(DView As DrawingView)
        Dim TopLeft = TransGeom.CreatePoint2d(DView.Left, DView.Top)
        Dim TopRight = TransGeom.CreatePoint2d(DView.Left + DView.Width, DView.Top)
        Dim BtmLeft = TransGeom.CreatePoint2d(DView.Left, DView.Top - DView.Height)
        Dim BtmRight = TransGeom.CreatePoint2d(DView.Left + DView.Width, DView.Top - DView.Height)

        _topLine = TransGeom.CreateLineSegment2d(TopLeft, TopRight)
        _btmLine = TransGeom.CreateLineSegment2d(BtmLeft, BtmRight)
        _leftLine = TransGeom.CreateLineSegment2d(TopLeft, BtmLeft)
        _rightLine = TransGeom.CreateLineSegment2d(TopRight, BtmRight)

        _blnViewMargin = 3
        _blnVerticalOffset = 2.0
        _blnHorizontalOffset = 2.0
    End Sub

    Private Sub CreateRowItemBalloon(Item As PartsListRow, View As DrawingView)
        Dim attachPoint As GeometryIntent = GetBalloonAttachGeometry(Item, View)
        If attachPoint Is Nothing Then Exit Sub

        Dim leaderPoint As Point2d = GetBalloonPosition(attachPoint.PointOnSheet, View)
        If leaderPoint Is Nothing Then Exit Sub

        Dim LeaderPoints As ObjectCollection = Oapp.TransientObjects.CreateObjectCollection
        LeaderPoints.Add(leaderPoint)
        LeaderPoints.Add(attachPoint)
        View.Parent.Balloons.Add(LeaderPoints)
    End Sub

    Private Sub ArrangeBalloonsAroundView(View As DrawingView)
    Dim allBalloons As New List(Of Balloon)
    For Each balloon As Balloon In View.Parent.Balloons
        If balloon.ParentView Is View Then
            allBalloons.Add(balloon)
        End If
    Next

    Dim topBalloons As New List(Of Balloon)
    Dim bottomBalloons As New List(Of Balloon)
    Dim leftBalloons As New List(Of Balloon)
    Dim rightBalloons As New List(Of Balloon)

    ' Classify balloons by approximate quadrant
    For Each balloon As Balloon In allBalloons
        Dim attachX = balloon.Position.X
        Dim attachY = balloon.Position.Y

        If attachY > View.Center.Y + View.Height / 4 Then
            topBalloons.Add(balloon)
        ElseIf attachY < View.Center.Y - View.Height / 4 Then
            bottomBalloons.Add(balloon)
        ElseIf attachX < View.Center.X Then
            leftBalloons.Add(balloon)
        Else
            rightBalloons.Add(balloon)
        End If
    Next

    ' Distribute using vertical and horizontal offsets
    DistributeBalloonsAlongTopOrBottom(topBalloons, View.Top + _blnViewMargin + _blnVerticalOffset, View.Left, View.Left + View.Width)
    DistributeBalloonsAlongTopOrBottom(bottomBalloons, View.Top - View.Height - _blnViewMargin - _blnVerticalOffset, View.Left, View.Left + View.Width)
    DistributeBalloonsAlongLeftOrRight(leftBalloons, View.Left - _blnViewMargin - _blnHorizontalOffset, View.Top, View.Top - View.Height)
    DistributeBalloonsAlongLeftOrRight(rightBalloons, View.Left + View.Width + _blnViewMargin + _blnHorizontalOffset, View.Top, View.Top - View.Height)
End Sub

    Private Sub DistributeBalloonsAlongTopOrBottom(balloons As List(Of Balloon), yPos As Double, xStart As Double, xEnd As Double)
        If balloons.Count = 0 Then Exit Sub

        balloons.Sort(Function(a, b) a.Position.X.CompareTo(b.Position.X))
        Dim spacing As Double = (xEnd - xStart) / (balloons.Count + 1)

        For i = 0 To balloons.Count - 1
            Dim newX = xStart + spacing * (i + 1)
            balloons(i).Position = TransGeom.CreatePoint2d(newX, yPos)
        Next
    End Sub

    Private Sub DistributeBalloonsAlongLeftOrRight(balloons As List(Of Balloon), xPos As Double, yStart As Double, yEnd As Double)
        If balloons.Count = 0 Then Exit Sub

        balloons.Sort(Function(a, b) a.Position.Y.CompareTo(b.Position.Y))
        Dim spacing As Double = (yStart - yEnd) / (balloons.Count + 1)

        For i = 0 To balloons.Count - 1
            Dim newY = yStart - spacing * (i + 1)
            balloons(i).Position = TransGeom.CreatePoint2d(xPos, newY)
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
            Case Else
                Return Nothing
        End Select
        Return leaderPoint
    End Function

    Private Function GetQuadrant(AttachPoint As Point2d, View As DrawingView) As String
        Dim CornerAngle As Double = Math.Atan2(View.Height, View.Width)
        Dim PointAngle As Double = Math.Atan2(AttachPoint.Y - View.Center.Y, AttachPoint.X - View.Center.X)
        If PointAngle < CornerAngle AndAlso PointAngle > -CornerAngle Then Return "Right"
        If PointAngle > CornerAngle AndAlso PointAngle < Math.PI - CornerAngle Then Return "Top"
        If PointAngle > Math.PI - CornerAngle OrElse PointAngle < -Math.PI + CornerAngle Then Return "Left"
        If PointAngle > -Math.PI + CornerAngle AndAlso PointAngle < -CornerAngle Then Return "Bottom"
        Return "Quadrant not Found"
    End Function

    Private Function GetBalloonAttachGeometry(Item As PartsListRow, View As DrawingView) As GeometryIntent
        Dim occEnum As ComponentOccurrencesEnumerator = View.ReferencedDocumentDescriptor.ReferencedDocument.ComponentDefinition.Occurrences.AllReferencedOccurrences(Item.ReferencedFiles.Item(1).DocumentDescriptor)
        Dim curves As List(Of DrawingCurve) = GetCurvesFromOcc(occEnum, View)
        Return GetAttachPoint(GetBestSegmentFromOccurrence(curves))
    End Function

#Region "Curve helpers"

    Private Function GetCurvesFromOcc(Occurrences As ComponentOccurrencesEnumerator, View As DrawingView) As List(Of DrawingCurve)
        Dim curves As New List(Of DrawingCurve)
        For Each occ As ComponentOccurrence In Occurrences
            If Not occ.Suppressed Then
                For Each curve As DrawingCurve In View.DrawingCurves(occ)
                    If Not curves.Contains(curve) Then curves.Add(curve)
                Next
            End If
        Next
        Return curves
    End Function

    Private Function GetBestSegmentFromOccurrence(curves As List(Of DrawingCurve)) As DrawingCurveSegment
        Dim bestRating As Double = 0
        Dim bestSegment As DrawingCurveSegment = Nothing
        For Each curve As DrawingCurve In curves
            Dim rating As Double
            Dim seg As DrawingCurveSegment = GetBestSegmentFromCurve(curve, rating)
            If rating > bestRating Then
                bestRating = rating
                bestSegment = seg
            End If
        Next
        Return bestSegment
    End Function

    Private Function GetBestSegmentFromCurve(curve As DrawingCurve, ByRef segmentRating As Double) As DrawingCurveSegment
        Dim bestSegment As DrawingCurveSegment = Nothing
        segmentRating = 0
        For Each segment As DrawingCurveSegment In curve.Segments
            Dim segLength As Double = GetSegmentLength(segment)
            If segLength > segmentRating Then
                bestSegment = segment
                segmentRating = segLength
            End If
        Next
        Return bestSegment
    End Function

    Private Function GetSegmentLength(segment As DrawingCurveSegment) As Double
        Select Case segment.GeometryType
            Case Curve2dTypeEnum.kLineCurve2d
                Return segment.StartPoint.DistanceTo(segment.EndPoint)
            Case Else
                Return GetCurveLength(segment.Geometry.Evaluator)
        End Select
    End Function

    Private Function GetCurveLength(eval As Curve2dEvaluator) As Double
        Dim minParam As Double, maxParam As Double, length As Double
        eval.GetParamExtents(minParam, maxParam)
        eval.GetLengthAtParam(minParam, maxParam, length)
        Return length
    End Function

    Private Function GetAttachPoint(segment As DrawingCurveSegment) As GeometryIntent
        If segment Is Nothing Then Return Nothing
        Dim sheet As Sheet = segment.Parent.Parent.Parent
        Return sheet.CreateGeometryIntent(segment.Parent, GetSegmentMidPoint(segment))
    End Function

    Private Function GetSegmentMidPoint(segment As DrawingCurveSegment) As Point2d
        Dim eval As Curve2dEvaluator = segment.Geometry.Evaluator
        Dim minParam, maxParam, midParam(0), point(1), length As Double
        eval.GetParamExtents(minParam, maxParam)
        eval.GetLengthAtParam(minParam, maxParam, length)
        eval.GetParamAtLength(minParam, length / 2, midParam(0))
        eval.GetPointAtParam(midParam, point)
        Return TransGeom.CreatePoint2d(point(0), point(1))
    End Function

#End Region

    ' Private properties
    Private Property _blnViewMargin As Double
    Private Property _blnVerticalOffset As Double
    Private Property _blnHorizontalOffset As Double
    Private Property _topLine As LineSegment2d
    Private Property _btmLine As LineSegment2d
    Private Property _rightLine As LineSegment2d
    Private Property _leftLine As LineSegment2d