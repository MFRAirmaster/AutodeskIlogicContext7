' Title: Sketched Symbol
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sketched-symbol/td-p/13813702#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:08:50.481109

Sub Main()
AddOrInsertPromptedSymbol()
InsertSymbolsNextToDimensions()
End Sub

Sub AddOrInsertPromptedSymbol()
Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument
Dim symbolName As String = "My Symbol"
Dim found As Boolean = False

' Check if symbol definition already exists
For Each oSymbolDef As SketchedSymbolDefinition In oDoc.SketchedSymbolDefinitions
If oSymbolDef.Name = symbolName Then
found = True
Exit For
End If
Next

If found Then Return

' Create the sketched symbol definition
Dim oSymbolDefNew As SketchedSymbolDefinition = oDoc.SketchedSymbolDefinitions.Add(symbolName)
Dim oSketch As DrawingSketch = Nothing
oSymbolDefNew.Edit(oSketch)

' Draw a circle and add prompted text box in the symbol sketch
Dim oCenter As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(12, 12)
oSketch.SketchCircles.AddByCenterRadius(oCenter, 0.25)
Dim promptPlaceholder As String = "<Prompt>Enter Text</Prompt>"
Dim oTextBox As TextBox = oSketch.TextBoxes.AddFitted(oCenter, promptPlaceholder)

oSymbolDefNew.ExitEdit(True)

' Do NOT insert the symbol instance here to avoid error on first run
End Sub

Sub InsertSymbolsNextToDimensions()
Dim oDoc As DrawingDocument = ThisDoc.Document
Dim oSymDef As SketchedSymbolDefinition

' Get the "BA" sketched symbol definition
Try
oSymDef = oDoc.SketchedSymbolDefinitions.Item("My Symbol")
Catch ex As Exception
MsgBox("Sketch Symbol 'BA' not found: " & ex.Message)
Exit Sub
End Try

Dim counter As Integer = 1
Dim offsetX As Double = 0.5
Dim offsetY As Double = 0.2
Dim oTG As TransientGeometry = ThisApplication.TransientGeometry

For Each oSheet As Sheet In oDoc.Sheets
Dim oDims As GeneralDimensions = oSheet.DrawingDimensions.GeneralDimensions
For Each oDim As GeneralDimension In oDims
Dim dimPt As Point2d = oDim.Text.Origin
Dim symPt As Point2d = oTG.CreatePoint2d(dimPt.X + offsetX, dimPt.Y + offsetY)
Dim oLeaderPoints As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection()

' Create geometry intent at the dimension location
Dim oIntent As GeometryIntent = oSheet.CreateGeometryIntent(oDim, dimPt)
oLeaderPoints.Add(symPt)
oLeaderPoints.Add(oIntent)

Dim promptValues(0) As String
promptValues(0) = counter.ToString()

Try
Dim oSketchedSymbol As SketchedSymbol = oSheet.SketchedSymbols.AddWithLeader(oSymDef, oLeaderPoints, 0, 1, promptValues)
oSketchedSymbol.Static = True
oSketchedSymbol.LeaderVisible = False
Catch ex As Exception
MsgBox("Error placing symbol " & counter & ": " & ex.Message)
End Try

counter += 1
Next
Next
End Sub