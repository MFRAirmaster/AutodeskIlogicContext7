' Title: Add sketchedsymbol and weldingsymbol
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/add-sketchedsymbol-and-weldingsymbol/td-p/13746582#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:05:13.877705

Public Function AddDrawingWeldingSymbol()

    ' Set a reference to the drawing document.
    ' This assumes a drawing document is active.
    Dim oDrwDoc As DrawingDocument
    Set oDrwDoc = ThisApplication.ActiveDocument

    ' Set a reference to the active sheet.
    Dim oActiveSheet As Sheet
    Set oActiveSheet = oDrwDoc.ActiveSheet

    ' Obtain a reference to the desired sketched symbol definition.
    Dim oSketchedSymbolDef As SketchedSymbolDefinition
    Set oSketchedSymbolDef = oDrwDoc.SketchedSymbolDefinitions.Item("Schweisshinweis_01")

    Dim getPoint As New CLS_GetPoint
    Dim pnt As Point2d
    
    Set pnt = getPoint.GetDrawingPoint("Platzierpunkt", kLeftMouseButton)
    
    ' Set a reference to the TransientGeometry object.
    Dim oTG As TransientGeometry
    Set oTG = ThisApplication.TransientGeometry
    
    Dim oPointPlace As Point2d
    Set oPointPlace = oTG.CreatePoint2d(pnt.X, pnt.Y)
    
    ' Add an instance of the sketched symbol definition to the sheet.
    Dim oSketchedSymbol As SketchedSymbol
    Set oSketchedSymbol = oActiveSheet.SketchedSymbols.Add(oSketchedSymbolDef, oPointPlace)

    Dim oLeaderPoints As ObjectCollection
    Set oLeaderPoints = ThisApplication.TransientObjects.CreateObjectCollection

    ' Create a few leader points.
    Call oLeaderPoints.Add(oPointPlace)
    
    Dim oLine As SketchLine
    Set oLine = oSketchedSymbol.Definition.Sketch.SketchLines(1)

    ' Create an intent and add to the leader points collection.
    ' This is the geometry that the symbol will attach to.
    Dim oGeometryIntent As GeometryIntent
    Set oGeometryIntent = oActiveSheet.CreateGeometryIntent(oLine, oLine.StartSketchPoint)

    Call oLeaderPoints.Add(oGeometryIntent)
    
    Dim oWeldingSymDefs As DrawingWeldingSymbolDefinitions
    Set oWeldingSymDefs = oActiveSheet.WeldingSymbols.CreateDefinitions()
   
    Dim oWeldingSymDef As DrawingWeldingSymbolDefinition
    Set oWeldingSymDef = oWeldingSymDefs.Add(1)
    
    ' Specify the weld symbol type(WeldSymbolTypeEnum/BackingSymbolTypeEnum)
    ' Kehlnaht
    oWeldingSymDef.WeldSymbolOne.WeldSymbolType = 124162
    oWeldingSymDef.WeldSymbolOne.SizeOrStrength = "a3"
    
    ' Create the symbol with a leader
    Dim oSymbol As DrawingWeldingSymbol
    Set oSymbol = oActiveSheet.WeldingSymbols.Add(oLeaderPoints, oWeldingSymDefs)
    
End Function