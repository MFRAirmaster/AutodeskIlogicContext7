' Title: Use Geometry Intent to place a balloon
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/use-geometry-intent-to-place-a-balloon/td-p/13724650#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:02:03.249536

Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

Dim WPt3 = VIEW1.GetIntent("Work Point2").PointOnSheet
WPt4  = ThisDrawing.Geometry.Point2d(WPt3.X * 10, WPt3.Y * 10)

Dim WPt = VIEW1.GetIntent("Work Point3").PointOnSheet
WPt2  = ThisDrawing.Geometry.Point2d(WPt.X * 10, WPt.Y * 10)

Dim balloon1 = Sheet_1.Balloons.Add("Balloon1", {WPt2}, {WPt4})