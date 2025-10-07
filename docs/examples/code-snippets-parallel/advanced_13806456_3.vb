' Title: Reference example: adding 2 leaders to a leader note
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/reference-example-adding-2-leaders-to-a-leader-note/td-p/13806456#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:35:46.245811

Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisApplication.ActiveDocument

' a reference to the active sheet.
Dim oActiveSheet As Sheet
oActiveSheet = oDrawDoc.ActiveSheet

' a reference to the drawing curve segment.
' Please select a linear drwaing curve.
Dim oDrawingCurveSegment1 As DrawingCurveSegment
oDrawingCurveSegment1 = ThisApplication.CommandManager.Pick(kDrawingCurveSegmentFilter, "Select a linear curve")

' a reference to the drawing curve.
Dim oDrawingCurve1 As DrawingCurve
oDrawingCurve1 = oDrawingCurveSegment1.Parent

' Get the mid point of the selected curve
' assuming that the selection curve is linear
Dim oMidPoint1 As Point2d
oMidPoint1 = oDrawingCurve1.MidPoint

Dim oDrawingCurveSegment2 As DrawingCurveSegment
oDrawingCurveSegment2 = ThisApplication.CommandManager.Pick(kDrawingCurveSegmentFilter, "Select a linear curve for the 2nd leader")

' a reference to the drawing curve.
Dim oDrawingCurve2 As DrawingCurve
oDrawingCurve2 = oDrawingCurveSegment2.Parent

' Get the mid point of the selected curve
' assuming that the selection curve is linear
Dim oMidPoint2 As Point2d
oMidPoint2 = oDrawingCurve2.MidPoint

 'a reference To the TransientGeometry Object.
Dim oTG As TransientGeometry
oTG = ThisApplication.TransientGeometry

Dim oLeaderPoints As ObjectCollection
oLeaderPoints = ThisApplication.TransientObjects.CreateObjectCollection

' Create a few leader points.
oLeaderPoints.Add(oTG.CreatePoint2d(oMidPoint1.X + 5, oMidPoint1.Y + 10))
'oLeaderPoints.Add(oTG.CreatePoint2d(oMidPoint1.X + 10, oMidPoint1.Y + 10)) ' use this to add a second point to the leader 

' Create an intent and add to the leader points collection.
' This is the geometry that the symbol will attach to.
Dim oGeometryIntent1 As GeometryIntent
oGeometryIntent1 = oActiveSheet.CreateGeometryIntent(oDrawingCurve1)

Dim oGeometryIntent2 As GeometryIntent
oGeometryIntent2 = oActiveSheet.CreateGeometryIntent(oDrawingCurve2)

Call oLeaderPoints.Add(oGeometryIntent1)

oMsg = InputBox("Enter text", "iLogic", "Hello World")

Dim oLeaderNote As LeaderNote
oLeaderNote = oActiveSheet.DrawingNotes.LeaderNotes.Add(oLeaderPoints,oMsg)

oLeaderPoints.Clear
Call oLeaderPoints.Add(oGeometryIntent2)
oLeaderNote.Leader.RootNode.AddLeader(oLeaderPoints)