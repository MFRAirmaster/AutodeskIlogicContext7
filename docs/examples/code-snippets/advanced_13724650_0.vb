' Title: Use Geometry Intent to place a balloon
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/use-geometry-intent-to-place-a-balloon/td-p/13724650
' Category: advanced
' Scraped: 2025-10-07T13:15:58.139127

Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

Dim oInt = VIEW1.GetIntent("FlyWheel", "Pie Face2" , intent := PointIntentEnum.kMidPointIntent)

Dim WPt = VIEW1.GetIntent( "Work Point1").PointOnSheet
WPt2  = ThisDrawing.Geometry.Point2d(WPt.X *.3937, WPt.Y*.3937)

Dim balloon1 = Sheet_1.Balloons.Add("Balloon1", {WPt2}, oInt)