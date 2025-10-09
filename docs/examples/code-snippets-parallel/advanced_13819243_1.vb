' Title: [iLogic/API] - Midpoint of dimension
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-api-midpoint-of-dimension/td-p/13819243
' Category: advanced
' Scraped: 2025-10-09T08:55:13.787083

' 1. Get the drawing document and active sheet
Dim oDoc As DrawingDocument = ThisDoc.Document
Dim oSheet As Sheet = oDoc.ActiveSheet

' 2. Prompt user to pick a dimension
Dim oDim As DrawingDimension = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingDimensionFilter, "Pick a Dimension")
If oDim Is Nothing Then Return

' 3. Get the midpoint of the dimension line
Dim oLine As LineSegment2d = oDim.DimensionLine
Dim midX As Double = (oLine.StartPoint.X + oLine.EndPoint.X) / 2
Dim midY As Double = (oLine.StartPoint.Y + oLine.EndPoint.Y) / 2

' 4. Define offset for symbol position (adjust as needed, units are usually cm or inches based on your template)
Dim offsetX As Double = 1.0  ' 1 cm to the right
Dim offsetY As Double = 0.5  ' 0.5 cm upward

' 5. Calculate the offset point for symbol insertion
Dim oOffsetPoint As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(midX + offsetX, midY + offsetY)

' 6. Get the sketched symbol definition name - replace with your actual symbol name
Dim symbolName As String = "AnynameofSketchsymbols"

' 7. Check if the sketched symbol exists in the document
Dim skDef As SketchedSymbolDefinition = Nothing
For Each def As SketchedSymbolDefinition In oDoc.SketchedSymbolDefinitions
    If def.Name = symbolName Then
        skDef = def
        Exit For
    End If
Next

If skDef Is Nothing Then
    MsgBox("Sketched symbol '" & symbolName & "' not found.", vbExclamation, "Error")
    Return
End If

' 8. Create geometry intent at the original dimension midpoint (no offset here)
Dim oGeometryIntent As GeometryIntent = oSheet.CreateGeometryIntent(oDim, ThisApplication.TransientGeometry.CreatePoint2d(midX, midY))

' 9. Prepare leader points: start at symbol offset point, end at geometry intent
Dim oLeaderPoints As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
oLeaderPoints.Add(oOffsetPoint)       ' Symbol insertion point (offset)
oLeaderPoints.Add(oGeometryIntent)    ' Leader target (dimension)

' 10. Prepare prompt strings for the symbol (empty if none)
Dim promptStrings() As String = {}

' 11. Add the sketched symbol with leader to the sheet
Dim skSymbol As SketchedSymbol = oSheet.SketchedSymbols.AddWithLeader(skDef, oLeaderPoints, 0, 1, promptStrings)
skSymbol.LeaderVisible = True
skSymbol.Static = True