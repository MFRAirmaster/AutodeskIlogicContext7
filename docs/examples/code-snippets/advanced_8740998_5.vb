' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-07T13:35:54.731250

Sub Main()
	
	sPattName = "BackerPattern"
	
	Dim oPatterns As OccurrencePatterns
	oPatterns = ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns
	
	Try
		oPatterns.Item(sPattName).Delete
	Catch ex As Exception
		'MsgBox("Error: " & ex.Message)
	End Try
	
	 
	Dim oPatt = Patterns.AddRectangular("BackerPattern", "Backer_01", _
	4, 12, Nothing, "X Axis", columnNaturalDirection := True, 
	rowCount := 3, rowOffset := 8, rowEntityName := "Y Axis", rowNaturalDirection := True)
	
	Dim oPattern As RectangularOccurrencePattern
	oPattern = ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns.Item(sPattName)
	oPattern.Name = sPattName
	
	oPattern.ColumnCount.Name = "BP_Col_Qty"
	oPattern.ColumnOffset.Name = "BP_Col_Spc"
	
	oPattern.RowCount.Name = "BP_Row_Qty"
	oPattern.RowOffset.Name = "BP_Row_Spc"
	
	ThisApplication.ActiveView.Fit
End Sub