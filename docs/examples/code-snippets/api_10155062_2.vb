' Title: Export Drawing dimension to excel
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062#messageview_0
' Category: api
' Scraped: 2025-10-07T13:37:26.659195

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.Save

	Next