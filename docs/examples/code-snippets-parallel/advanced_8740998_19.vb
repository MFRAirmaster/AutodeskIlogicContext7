' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-09T09:06:46.157662

Patterns.AddRectangular("BLADE PATTERN",
	{"BLADE", "BACK-RIGHT BLADE SCREW", "FRONT-RIGHT BLADE SCREW", "BACK-LEFT BLADE SCREW", "FRONT-LEFT BLADE SCREW" },
	Parameter("NO_OF_BLADES"), Parameter("Sb"), "", "Y Axis")

	Dim oPattern As RectangularOccurrencePattern
	'Find pattern through API
	oPattern = ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns.Item("BLADE PATTERN")
	'Change the Column count expression to a parameter
	oPattern.ColumnCount.Expression = Parameter.Param("NO_OF_BLADES").Name
	'Change the column offset expression to a parameter
	oPattern.ColumnOffset.Expression = Parameter.Param("Sb").Name