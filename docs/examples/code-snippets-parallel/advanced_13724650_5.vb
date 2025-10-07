' Title: Use Geometry Intent to place a balloon
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/use-geometry-intent-to-place-a-balloon/td-p/13724650#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:11:46.238617

Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

Dim Edge0 = VIEW1.GetIntent("Frame Base:1", "Edge0" , intent := PointIntentEnum.kMidPointIntent)

Dim WPt1 = VIEW1.GetIntent("Work Point3").PointOnSheet
WPt3  = ThisDrawing.Geometry.Point2d(WPt1.X * 10, WPt1.Y * 10)
Dim WPt2 = VIEW1.GetIntent("Work Point4").PointOnSheet
WPt4  = ThisDrawing.Geometry.Point2d(WPt2.X * 10 , WPt2.Y * 10)

Dim balloon1 = Sheet_1.Balloons.Add("Balloon1", {WPt4, WPt3}, Edge0)