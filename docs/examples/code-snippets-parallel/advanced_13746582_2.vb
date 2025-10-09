' Title: Add sketchedsymbol and weldingsymbol
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/add-sketchedsymbol-and-weldingsymbol/td-p/13746582
' Category: advanced
' Scraped: 2025-10-09T08:58:36.120898

Public Function InsertSketchedSymbolOnSheet()
    ' Set a reference to the drawing document.
    ' This assumes a drawing document is active.
    Dim oDrwDoc As DrawingDocument
    Set oDrwDoc = ThisApplication.ActiveDocument

'    ' Obtain a reference to the desired sketched symbol definition.
'    Dim oSSDefs As SketchedSymbolDefinitions
'    Set oSSDefs = oDrwDoc.SketchedSymbolDefinitions
'
'    Dim oLibrary As SketchedSymbolDefinitionLibrary
'    Set oLibrary = oSSDefs.SketchedSymbolDefinitionLibraries.Item("ROESLER-symbol-library-01.idw")
'
'    Dim oSketchedSymbolDef As SketchedSymbolDefinition
'    Set oSketchedSymbolDef = oSSDefs.AddFromLibrary(oLibrary, "Schweißhinweis", True)
    
    ' Obtain a reference to the desired sketched symbol definition.
    Dim oSketchedSymbolDef As SketchedSymbolDefinition
    Set oSketchedSymbolDef = oDrwDoc.SketchedSymbolDefinitions.Item("Schweißhinweis")
    
    Dim oSheet As Sheet
    Set oSheet = oDrwDoc.ActiveSheet

    Dim getPoint As New CLS_GetPoint
    Dim pnt As Point2d
    
    Set pnt = getPoint.GetDrawingPoint("Click the desired location", kLeftMouseButton)

    Dim oTG As TransientGeometry
    Set oTG = ThisApplication.TransientGeometry
    
    Dim oPoint1 As Point2d
    Set oPoint1 = oTG.CreatePoint2d(pnt.X, pnt.Y)
    
    Dim oLine As SketchLine
    Set oLine = oSketchedSymbolDef.Sketch.SketchLines(1)


    ' Add an instance of the sketched symbol definition to the sheet.
    Dim oSketchedSymbol As SketchedSymbol
    Set oSketchedSymbol = oSheet.SketchedSymbols.Add(oSketchedSymbolDef, oPoint1)
    
    Dim oLeaderPoints As ObjectCollection
    Set oLeaderPoints = ThisApplication.TransientObjects.CreateObjectCollection

    ' Create a few leader points.
    Call oLeaderPoints.Add(oPoint1)

    ' Create an intent and add to the leader points collection.
    ' This is the geometry that the symbol will attach to.
    Dim oGeometryIntent As GeometryIntent
    Set oGeometryIntent = oSheet.CreateGeometryIntent(oLine)

    Call oLeaderPoints.Add(oGeometryIntent)
    
    Dim oWeldingSymDefs As DrawingWeldingSymbolDefinitions
    Set oWeldingSymDefs = oSheet.WeldingSymbols.CreateDefinitions()
   
    Dim oWeldingSymDef As DrawingWeldingSymbolDefinition
    Set oWeldingSymDef = oWeldingSymDefs.Add(1)
     
    ' Specify the weld symbol type(WeldSymbolTypeEnum/BackingSymbolTypeEnum)
    ' Kehlnaht
    oWeldingSymDef.WeldSymbolOne.WeldSymbolType = 124162
    oWeldingSymDef.WeldSymbolOne.SizeOrStrength = "a3"
    
    ' Create the symbol with a leader
    Dim oSymbol As DrawingWeldingSymbol
    Set oSymbol = oSheet.WeldingSymbols.Add(oLeaderPoints, oWeldingSymDefs)
    
End Function