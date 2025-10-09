' Title: Automated Break on Front View
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automated-break-on-front-view/td-p/13765786#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:51:08.166712

Public Sub CreateBreakoperationInDrawingView()
    ' Set a reference to the drawing document.
    ' This assumes a drawing document is active.
    Dim oDrawDoc As DrawingDocument = ThisApplication.ActiveDocument

    'Set a reference to the active sheet.
    Dim oSheet As Sheet = oDrawDoc.ActiveSheet

    ' Check to make sure a drawing view is selected.
    If Not TypeOf oDrawDoc.SelectSet.Item(1) Is DrawingView Then
        MsgBox "A drawing view must be selected."
        Exit Sub
    End If

    ' Set a reference to the selected drawing. This assumes
    ' that the selected view is not a draft view.
    Dim oDrawingView As DrawingView = oDrawDoc.SelectSet.Item(1)

    ' Set a reference to the center of the base view.
    Dim oCenter As Point2d = oDrawingView.Center

    ' Define the start point of the break
    Dim oStartPoint As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(oCenter.X - 1, oCenter.Y)

    ' Define the end point of the break
    Dim oEndPoint As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(oCenter.X + 1, oCenter.Y)

    Dim oBreakOperation As BreakOperation = oDrawingView.BreakOperations.Add(kHorizontalBreakOrientation, oStartPoint, oEndPoint, kRectangularBreakStyle, 5)
End Sub