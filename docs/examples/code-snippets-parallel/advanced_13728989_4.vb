' Title: Using Named Entities to Automate Dimension using Inventor API.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/using-named-entities-to-automate-dimension-using-inventor-api/td-p/13728989
' Category: advanced
' Scraped: 2025-10-07T14:05:16.543015

Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

Dim namedGeometry1 = VIEW1.GetIntent("BackFace")
Dim namedGeometry2 = VIEW1.GetIntent("FrontFace")

Dim genDims = Sheet_1.DrawingDimensions.GeneralDimensions

ThisDrawing.BeginManage()
If True Then 
	Dim linDim1 = genDims.AddLinear("Dimension 1", VIEW1.SheetPoint(1.1, 0.5), namedGeometry1)
	Dim linDim2 = genDims.AddLinear("Dimension 2", VIEW1.SheetPoint(-0.1, 0.5), namedGeometry2)
	Dim linDim3 = genDims.AddLinear("Dimension 3", VIEW1.SheetPoint(0.5, 1.2), namedGeometry1, namedGeometry2)
End If
ThisDrawing.EndManage()