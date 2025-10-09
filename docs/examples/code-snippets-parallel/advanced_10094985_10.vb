' Title: Drawing View Labels
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/drawing-view-labels/td-p/10094985
' Category: advanced
' Scraped: 2025-10-09T09:08:53.049118

Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
If oDDoc Is Nothing Then Return
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oDView As DrawingView In oSheet.DrawingViews
		If (TypeOf oDView Is DetailDrawingView) Then Continue For
		If (TypeOf oDView Is SectionDrawingView) Then Continue For
		Dim sViewOrientType As String = oDView.Camera.ViewOrientationType.ToString()
		oDView.ShowLabel = True
		oDView.Name = Mid(sViewOrientType, 2, Len(sViewOrientType) -16).ToUpper()
	Next
Next