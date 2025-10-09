' Title: Export Drawing dimension to excel
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062
' Category: api
' Scraped: 2025-10-09T09:04:30.935192

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	'MsgBox (b.Tolerance.Upper)
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	GoExcel.Save

	Next