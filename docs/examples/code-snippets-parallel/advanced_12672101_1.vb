' Title: Get overall dimensions  of all views of a drawing by iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/get-overall-dimensions-of-all-views-of-a-drawing-by-ilogic/td-p/12672101
' Category: advanced
' Scraped: 2025-10-07T14:35:53.111302

Public Class ThisRule
' This code was inspired and uses much of a code by Jelte de Jong

    Private _doc As DrawingDocument
    Private _sheet As Sheet
    Private _view As DrawingView
Private _intents As List(Of GeometryIntent) = New List(Of GeometryIntent)()

Sub main
' Check a document is open
If ThisApplication.Documents.VisibleDocuments.Count = 0 Then
	MessageBox.Show("A Drawing Document needs to be open")
	Exit Sub
End If

' Check active document
If ThisApplication.ActiveDocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
	MessageBox.Show("A Drawing Document needs to be open")
	Exit Sub
End If

' Get the active document
'Dim oDoc As Inventor.DrawingDocument = ThisApplication.ActiveDocument
_doc = ThisDoc.Document


' Iterate through all sheets
For Each _sheet As Inventor.Sheet In _doc.Sheets
	
	' Iterate through all views
	For Each _view As Inventor.DrawingView In _sheet.DrawingViews
		
		' Check view is a base view, skip if not
		If _view.ViewType <> Inventor.DrawingViewTypeEnum.kStandardDrawingViewType Then Continue For
		
		    CreateIntentList()
      		createHorizontalOuterDimension()
        	createVerticalOuterDimension()
		
		
	Next ' Next view
Next ' Next sheet
End Sub

'Public Class ThisRule
'	' This code was written by Jelte de Jong
'	' and published on www.hjalte.nl
'    Private _doc As DrawingDocument
'    Private _sheet As Sheet
'    Private _view As DrawingView
'    Private _intents As List(Of GeometryIntent) = New List(Of GeometryIntent)()

'    Sub Main()
'        _doc = ThisDoc.Document
'        _sheet = _doc.ActiveSheet
'        _view = ThisApplication.CommandManager.Pick(
'                       SelectionFilterEnum.kDrawingViewFilter, 
'                       "Select a drawing view")

'        CreateIntentList()
'        createHorizontalOuterDimension()
'        createVerticalOuterDimension()
'    End Sub

    Private Sub createHorizontalOuterDimension()
        Dim orderedIntents = _intents.OrderByDescending(Function(s) s.PointOnSheet.X)

        Dim pointLeft = orderedIntents.First
        Dim pointRight = orderedIntents.Last

        Dim textX = pointLeft.PointOnSheet.X +
                (pointRight.PointOnSheet.X - pointLeft.PointOnSheet.X) / 2
        Dim textY = _view.Position.Y - _view.Height / 2 - 2

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
        Dim textX = _view.Position.X + _view.Width / 2 + 2

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

End Class