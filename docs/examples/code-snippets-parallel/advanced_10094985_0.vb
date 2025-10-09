' Title: Drawing View Labels
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/drawing-view-labels/td-p/10094985
' Category: advanced
' Scraped: 2025-10-09T09:08:53.049118

Dim tmpSheet As Sheet = ThisDrawing.ActiveSheet.Sheet
Dim tmpView As DrawingView
For Each tmpView In tmpSheet.DrawingViews
    Dim myDrawingViewName As String = ""
	If tmpView.Type = ObjectTypeEnum.kDrawingViewObject Then
    
        Select Case tmpView.Camera.ViewOrientationType
            Case ViewOrientationTypeEnum.kBackViewOrientation
                myDrawingViewName = "BACK VIEW"
            Case ViewOrientationTypeEnum.kBottomViewOrientation
                myDrawingViewName = "BOTTOM VIEW"
            Case ViewOrientationTypeEnum.kFrontViewOrientation
                myDrawingViewName = "FRONT VIEW"
            Case ViewOrientationTypeEnum.kIsoBottomLeftViewOrientation
                myDrawingViewName = "ISOMETRIC VIEW"
            Case ViewOrientationTypeEnum.kIsoBottomRightViewOrientation
                myDrawingViewName = "ISOMETRIC VIEW"
            Case ViewOrientationTypeEnum.kIsoTopLeftViewOrientation
                myDrawingViewName = "ISOMETRIC VIEW"
            Case ViewOrientationTypeEnum.kIsoTopRightViewOrientation
                myDrawingViewName = "ISOMETRIC VIEW"
            Case ViewOrientationTypeEnum.kLeftViewOrientation
                myDrawingViewName = "SIDE VIEW"
            Case ViewOrientationTypeEnum.kRightViewOrientation
                myDrawingViewName = "SIDE VIEW"
            Case ViewOrientationTypeEnum.kTopViewOrientation
                myDrawingViewName = "TOP VIEW"
            Case Else
                myDrawingViewName = ""
            End Select        
        If Not myDrawingViewName = "" Then
            tmpView.Name = myDrawingViewName
        End If		
		If tmpView.IsFlatPatternView =True Then
           tmpView.Name = "FLAT PATTERN"
        End If
    End If
    myDrawingViewName = ""

Next