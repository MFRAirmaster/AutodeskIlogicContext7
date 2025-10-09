' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-09T09:06:46.157662

Dim rectPattern = Patterns.AddRectangular( _
	"RectPattern2", "WR157_01", _
	ColCnt, ColSpc, Nothing, "X Axis", columnNaturalDirection := True, _
	rowCount := RowCnt, rowOffset := RowSpc, rowEntityName := "Y Axis", rowNaturalDirection := True)