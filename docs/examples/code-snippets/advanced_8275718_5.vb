' Title: Auto Ordinate Dimensions
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-ordinate-dimensions/td-p/8275718#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:45:52.795768

Public Sub Main()
    ' Set a reference to the drawing document.
    ' This assumes a drawing document is active.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    
    ' Set a reference to the active sheet.
    Dim oActiveSheet As Sheet
    Set oActiveSheet = oDrawDoc.ActiveSheet
    
    Dim oDoc As Document
    Set oDoc = oActiveSheet.DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedDocument
    
    If oDoc.DocumentType = kAssemblyDocumentObject Then
    
        '   a reference to the drawing curve segment.
        ' This assumes that a linear drawing curve is selected.
        Dim oDrawingCurveSegment As DrawingCurveSegment
        Set oDrawingCurveSegment = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select start line to find occurrence")
        
        '   a reference to the drawing curve.
        Dim oDrawingCurve As DrawingCurve
        Set oDrawingCurve = oDrawingCurveSegment.Parent
        
        Dim oReferDef As AssemblyComponentDefinition
        Set oReferDef = oDoc.ComponentDefinition
        
        Dim pt As WorkPoint
        Set pt = oReferDef.WorkPoints.Item(1)
        
        Call oActiveSheet.DrawingViews.Item(1).SetIncludeStatus(pt, True)
        
        If Not oDrawingCurve.CurveType = CurveTypeEnum.kLineSegmentCurve Then
            MsgBox ("A linear curve should be selected for this sample.")
            Exit Sub
        End If
        
        Dim oCenterMark As Centermark
        Set oCenterMark = oActiveSheet.Centermarks.Item(oActiveSheet.Centermarks.Count)
        ' Create point intents to anchor the dimension to.
        Dim oDimIntent  As GeometryIntent
        Set oDimIntent = oActiveSheet.CreateGeometryIntent(oCenterMark, kCenterPointIntent)
        
        '   a reference to the view to which the curve belongs.
        Dim oDrawingView As DrawingView
        Set oDrawingView = oDrawingCurve.Parent
        
        ' If origin indicator has not been already created, create it first.
        If Not oDrawingView.HasOriginIndicator Then
            ' The indicator will be located at the start point of the selected curve.
            Call oDrawingView.CreateOriginIndicator(oDimIntent)
        End If
        
        '   a reference to the ordinate dimensions collection.
        Dim oOrdinateDimensions As OrdinateDimensions
        Set oOrdinateDimensions = oActiveSheet.DrawingDimensions.OrdinateDimensions
        
        Dim oTextOrigin  As Point2d
        Dim DimType As DimensionTypeEnum
        
        ' Selected curve is vertical or at an angle.
        DimType = DimensionTypeEnum.kVerticalDimensionType
        
        '   the text points for the 2 dimensions.
        
        Set oTextOrigin = ThisApplication.TransientGeometry.CreatePoint2d(oCenterMark.Position.X, oCenterMark.Position.Y - 2)
        
        ' Create the first ordinate dimension.
        Dim oOrdinateDimension1 As OrdinateDimension
        Set oOrdinateDimension1 = oOrdinateDimensions.Add(oDimIntent, oTextOrigin, DimType)
        
        Dim LArray() As String
        LArray = Split(oDrawingCurve.ModelGeometry.ContainingOccurrence.Name, ":")
        
        Dim occ As ComponentOccurrence
        For Each occ In oReferDef.Occurrences
            Dim occName() As String
            occName = Split(occ.Name, ":")
            If occName(0) = LArray(0) Then
                Set pt = occ.Definition.WorkPoints.Item(1)
                
                Dim oProxy As WorkPointProxy
                Call occ.CreateGeometryProxy(pt, oProxy)
                
                Call oActiveSheet.DrawingViews.Item(1).SetIncludeStatus(oProxy, True)
                
                Set oCenterMark = oActiveSheet.Centermarks.Item(oActiveSheet.Centermarks.Count)
                
                Dim oIntent As GeometryIntent
                Set oIntent = oActiveSheet.CreateGeometryIntent(oCenterMark, kCenterPointIntent)
                
                Dim origin As Point2d
                Set origin = ThisApplication.TransientGeometry.CreatePoint2d(oCenterMark.Position.X, oCenterMark.Position.Y - 2)
                
                Call oOrdinateDimensions.Add(oIntent, origin, DimType)
            End If
        Next
         
    
    Else
        MsgBox ("Referenced document is not an assembly document")
    End If
  
End Sub