' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-09T09:06:46.157662

Dim oPatt = Patterns.AddRectangular("BackerPattern", "Backer_01", BP_Col_Qty, BP_Col_Spc, Nothing, "X Axis", columnNaturalDirection := True, rowCount := BP_Row_Qty, rowOffset := BP_Row_Spc, rowEntityName := "Y Axis", rowNaturalDirection := True)