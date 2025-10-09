' Title: Creating a dimension from edge of circle with a face
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/creating-a-dimension-from-edge-of-circle-with-a-face/td-p/11569540
' Category: troubleshooting
' Scraped: 2025-10-09T08:53:48.928413

' Get sheet and view
Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

' Get named entities
Dim VFB_Face_Left = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Left", PointIntentEnum.kEndPointIntent)
Dim VFB_Face_Bottom = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Bottom", PointIntentEnum.kEndPointIntent)
Dim VRB_Face_Cylinder_Right = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularRightPointIntent)
Dim VRB_Face_Cylinder_Top = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularTopPointIntent)

' Declare access to drawing dimensions
Dim genDims = Sheet_1.DrawingDimensions.GeneralDimensions

' Add the dimensions
Try
	genDims.AddLinear("Dim01", VIEW1.SheetPoint(0.5, -0.1875), VFB_Face_Left, VRB_Face_Cylinder_Right, DimensionTypeEnum.kHorizontalDimensionType)
	genDims.AddLinear("Dim02", VIEW1.SheetPoint(-0.03125, 0), VFB_Face_Bottom, VRB_Face_Cylinder_Top, DimensionTypeEnum.kVerticalDimensionType)
Catch ex As Exception
	MsgBox("Error: ", ex.Message )
End Try