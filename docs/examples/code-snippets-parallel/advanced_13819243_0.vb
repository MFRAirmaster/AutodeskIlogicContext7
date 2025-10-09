' Title: [iLogic/API] - Midpoint of dimension
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-api-midpoint-of-dimension/td-p/13819243
' Category: advanced
' Scraped: 2025-10-09T08:55:13.787083

Dim oDoc As DrawingDocument = ThisDoc.Document
Dim oSheet As Sheet = oDoc.ActiveSheet

' 1. Pick Dimension
Dim oDim As DrawingDimension
oDim = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingDimensionFilter, "Pick Dimension")
If oDim Is Nothing Then Return

' 2. Get Midpoint =Problem here (This is a temporary solution because it doesn’t maintain a link with the dimension)
Dim oLine As LineSegment2d = oDim.DimensionLine
Dim midX As Double = (oLine.StartPoint.X + oLine.EndPoint.X) / 2
Dim midY As Double = (oLine.StartPoint.Y + oLine.EndPoint.Y) / 2
Dim oMid As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(midX, midY)

' 3. Insert Sketch Symbol
Dim skDef As SketchedSymbolDefinition = oDoc.SketchedSymbolDefinitions.Item("AnynameofSketchsymbols")
oSheet.SketchedSymbols.Add(skDef, oMid, 0)