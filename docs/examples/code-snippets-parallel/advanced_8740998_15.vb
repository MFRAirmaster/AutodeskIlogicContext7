' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-09T09:06:46.157662

ToPatternName = "TopLinks"
ToPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item("TopLinks")
ToPattern.Name = ToPatternName
Test5 = ToPattern.ColumnCount
MessageBox.Show(Test5,"Test 5")