' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-07T13:35:54.731250

Dim oDoc As AssemblyDocument = ThisDoc.Document
ToPatternName = "TopLinks"
Dim ToPattern As RectangularOccurrencePattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(ToPatternName)
Dim Test5 As Integer = ToPattern.ColumnCount.Value
MessageBox.Show(Test5,"Test 5")