' Title: specifiying an ordinate zero point that is on the geometry not a center mark
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/specifiying-an-ordinate-zero-point-that-is-on-the-geometry-not-a/td-p/13801442
' Category: advanced
' Scraped: 2025-10-07T12:28:27.385424

Sub Main
    Dim oDDoc As DrawingDocument = ThisDrawing.Document
    Dim oSheet As Sheet = oDDoc.ActiveSheet
    Dim oOrdDimSets As OrdinateDimensionSets = oSheet.DrawingDimensions.OrdinateDimensionSets
    Dim oOrdDimSet As OrdinateDimensionSet
    Dim oIntents As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
    Dim oTG As TransientGeometry = ThisApplication.TransientGeometry

    ' Collect center mark geometry intents
    For Each oCM As Centermark In oSheet.Centermarks
        Dim oGIntent As GeometryIntent = Nothing
        If oCM.Attached Then
            If TypeOf oCM.AttachedEntity Is GeometryIntent Then
                oGIntent = oCM.AttachedEntity
            End If
        End If
        If oGIntent Is Nothing Then
            oGIntent = oSheet.CreateGeometryIntent(oCM, PointIntentEnum.kCenterPointIntent)
        End If
        oIntents.Add(oGIntent)
    Next

    If oIntents.Count = 0 Then
        MsgBox("There were no 'Geometry Intents' added to the collection. Exiting.", , "")
        Exit Sub
    End If
	
' === ORDINATE: origin at left edge, baseline just outside the view ===
' Find BASE VIEW 
Dim BASE_VIEW As DrawingView = Nothing
For Each v As DrawingView In oSheet.DrawingViews
    If StrComp(v.Name, "BASE VIEW", vbTextCompare) = 0 Then
        BASE_VIEW = v : Exit For
    End If
Next
If BASE_VIEW Is Nothing Then
    MsgBox("BASE VIEW not found.", vbExclamation, "Ordinate")
    Exit Sub
End If

' Pick the LEFTMOST centermark as the origin 
Dim originGI As GeometryIntent = Nothing
Dim minX As Double = Double.MaxValue
For Each oCM As Centermark In oSheet.Centermarks
    Dim gi As GeometryIntent = oSheet.CreateGeometryIntent(oCM, PointIntentEnum.kCenterPointIntent)
    Dim pt As Point2d = gi.PointOnSheet
    If pt.X < minX Then
        minX = pt.X
        originGI = gi
    End If
Next
If originGI Is Nothing Then
    MsgBox("No centermarks found for ordinate set.", vbExclamation, "Ordinate")
    Exit Sub
End If

' Rebuild intents so ORIGIN is first
Dim intentsOrdered As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
intentsOrdered.Add(originGI)
For Each o In oIntents
    If TypeOf o Is GeometryIntent Then
        Dim gi = DirectCast(o, GeometryIntent)
        If Not (gi Is originGI) Then intentsOrdered.Add(gi)
    Else
        intentsOrdered.Add(o)
    End If
Next

' Baseline: vertical at LEFT edge 
Dim gap As Double = 0.00
Dim leftX As Double = BASE_VIEW.Center.X - (BASE_VIEW.Width / 2) - gap
Dim midY  As Double = BASE_VIEW.Center.Y
Dim oPlacement As Point2d = oTG.CreatePoint2d(leftX, midY)

' Create ordinate set
Try
    oOrdDimSet = oOrdDimSets.Add(intentsOrdered, oPlacement, DimensionTypeEnum.kHorizontalDimensionType)
Catch ex As Exception
    MsgBox("Failed to add vertical ordinate set." & vbCrLf & ex.Message, vbExclamation, "Ordinate")
    Exit Sub
End Try





End Sub