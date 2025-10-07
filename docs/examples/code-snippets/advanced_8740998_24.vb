' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-07T13:35:54.731250

Dim oPickedOcc As ComponentOccurrence = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select a Component.")
If oPickedOcc Is Nothing Then Return
Dim oADef As AssemblyComponentDefinition = oPickedOcc.Parent
Dim oOccPatts As OccurrencePatterns = oADef.OccurrencePatterns
Dim oColl As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
oColl.Add(oPickedOcc)
Dim oNewOccPatt As RectangularOccurrencePattern = _
oOccPatts.AddRectangularPattern(oColl, oADef.WorkAxes.Item(1), True, _
"mColumnOffset = ColumnOffset", "mColumnCount = ColumnCount", oADef.WorkAxes.Item(2), True, _
"mRowOffset = RowOffset", "mRowCount = RowCount")
oNewOccPatt.Name = "Rect Pattern Of " & oPickedOcc.Name.Split(":").First