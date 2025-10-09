' Title: Export Drawing dimension to excel
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062
' Category: api
' Scraped: 2025-10-09T09:04:30.935192

AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add'.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

'oCells.item(2,1).value= b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'MsgBox (b.Text.Text & " " & b.Tolerance.ToleranceType)
	
	If b.Tolerance.ToleranceType = 31236 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31233 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31235 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = "+/-" & b.Tolerance.Upper*10
		'oCells.item(i,3).value  = "+/-" &b.Tolerance.Lower*10.ToString
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If


	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close