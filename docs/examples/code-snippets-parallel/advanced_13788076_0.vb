' Title: Sketch symbol &amp; Table creation for QC Inspection
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sketch-symbol-amp-table-creation-for-qc-inspection/td-p/13788076#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:56:09.558894

Sub Main()
    Dim oDoc As DrawingDocument = ThisDoc.Document
    Dim oSymDef As SketchedSymbolDefinition

    Try
        oSymDef = oDoc.SketchedSymbolDefinitions.Item("Your Sketched Symbol Name")
    Catch ex As Exception
        MsgBox("Sketch Symbol not found: " & ex.Message)
        Exit Sub
    End Try

    Dim counter As Integer = 1
    Dim offsetX As Double = 0.5
    Dim offsetY As Double = 0.2
    Dim oTG As TransientGeometry = ThisApplication.TransientGeometry

    Dim oSheet As Sheet
    For Each oSheet In oDoc.Sheets
        Dim oDims As GeneralDimensions = oSheet.DrawingDimensions.GeneralDimensions

        For Each oDim As GeneralDimension In oDims
            Dim dimPt As Point2d = oDim.Text.Origin

            Dim symPt As Point2d = oTG.CreatePoint2d(dimPt.X + offsetX, dimPt.Y + offsetY)

            Dim oLeaderPoints As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection()
            oLeaderPoints.Add(symPt)
            oLeaderPoints.Add(dimPt)

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