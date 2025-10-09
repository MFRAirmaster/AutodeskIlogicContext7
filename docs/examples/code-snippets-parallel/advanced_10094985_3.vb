' Title: Drawing View Labels
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/drawing-view-labels/td-p/10094985
' Category: advanced
' Scraped: 2025-10-09T09:08:53.049118

Dim oDoc As DrawingDocument = ThisDrawing.Document
Dim oViews As DrawingViews
Dim oOrientType As ViewOrientationTypeEnum
Dim oOrigViewName As String
Dim oNewViewName As String
For Each oSheet As Sheet In oDoc.Sheets
	oViews = oSheet.DrawingViews
	For Each oView As DrawingView In oViews
		oOrientType = oView.Camera.ViewOrientationType
		oOrigViewName = oView.Name
		oNewViewName = Mid(oOrientType.ToString, 2, Len(oOrientType.ToString) -16)
		oView.ShowLabel = True
		oView.Name = oNewViewName.ToUpper
	Next
Next