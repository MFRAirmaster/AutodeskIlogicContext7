' Title: Unable to cast COM object of type 'System._ComObject to interface type 'Inventor.Point2d'.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/unable-to-cast-com-object-of-type-system-comobject-to-interface/td-p/13739507
' Category: advanced
' Scraped: 2025-10-07T13:54:49.288231

Imports System.Reflection

Sub Main()
Dim oDoc As Inventor.Document = ThisApplication.ActiveDocument
If oDoc.DocumentType <> Inventor.DocumentTypeEnum.kPartDocumentObject Then
    MessageBox.Show("Run this rule from a Part (.ipt) document.", "Error")
    Return
End If

' Create new drawing document
Dim oDrawingDoc As Inventor.DrawingDocument
' Use a specific template if you have one, otherwise default
' oDrawingDoc = ThisApplication.Documents.Add("C:\Users\Public\Documents\Autodesk\Inventor 202x\Templates\Metric\ISO.idw", True) ' Example for specific template
oDrawingDoc = ThisApplication.Documents.Add(Inventor.DocumentTypeEnum.kDrawingDocumentObject, _
    ThisApplication.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kDrawingDocumentObject))

Dim oSheet As Inventor.Sheet = oDrawingDoc.Sheets.Item(1)
Dim tg As Inventor.TransientGeometry = ThisApplication.TransientGeometry

' Adjust scale as needed based on part size and sheet size
Dim scale As Double = 0.1 ' Start with a reasonable scale, you might need to adjust

' Define view positions on the sheet. These are arbitrary and need fine-tuning.
' Units are in cm for a default A3 sheet (e.g., ISO.idw)
Dim basePoint As Inventor.Point2d = tg.CreatePoint2d(10, 20) ' Front view
Dim projRight As Inventor.Point2d = tg.CreatePoint2d(25, 20) ' Right view
Dim projIso As Inventor.Point2d = tg.CreatePoint2d(35, 25) ' Isometric view

Dim baseView As Inventor.DrawingView
Try
    baseView = oSheet.DrawingViews.AddBaseView(oDoc, basePoint, scale, _
        Inventor.ViewOrientationTypeEnum.kFrontViewOrientation, Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
Catch ex As Exception
    MessageBox.Show("Error adding base view: " & ex.Message, "View Creation Error")
    Return ' Exit if base view cannot be added
End Try

Dim rightView As Inventor.DrawingView
Try
    rightView = oSheet.DrawingViews.AddProjectedView(baseView, projRight, Inventor.DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
Catch ex As Exception
    MessageBox.Show("Error adding right view: " & ex.Message, "View Creation Error")
End Try

Dim isoView As Inventor.DrawingView
Try
    isoView = oSheet.DrawingViews.AddProjectedView(baseView, projIso, Inventor.DrawingViewStyleEnum.kShadedDrawingViewStyle)
Catch ex As Exception
    MessageBox.Show("Error adding isometric view: " & ex.Message, "View Creation Error")
End Try


' ===========================================
' Collect edges and points for dimensioning from the base view
' ===========================================
Dim oCurves As Inventor.DrawingCurvesEnumerator = baseView.DrawingCurves
Dim horizontalEdges As New List(Of Inventor.DrawingCurve) ' Store curves directly
Dim verticalEdges As New List(Of Inventor.DrawingCurve)   ' Store curves directly
Dim chamferCurves As New List(Of Inventor.DrawingCurve)
Dim circleCurves As New List(Of Inventor.DrawingCurve)

For Each c As Inventor.DrawingCurve In oCurves
    If c.Segments.Count > 0 Then
        If TypeOf c.Segments.Item(1).Geometry Is Inventor.LineSegment2d Then
            Dim lineSeg As Inventor.LineSegment2d = c.Segments.Item(1).Geometry
            Dim startPt As Inventor.Point2d = lineSeg.StartPoint
            Dim endPt As Inventor.Point2d = lineSeg.EndPoint

            If Math.Abs(startPt.Y - endPt.Y) < 0.01 Then ' Horizontal line
                horizontalEdges.Add(c)
            ElseIf Math.Abs(startPt.X - endPt.X) < 0.01 Then ' Vertical line
                verticalEdges.Add(c)
            Else ' Angled line, potential chamfer or other feature
                chamferCurves.Add(c)
            End If
        ElseIf c.Segments.Item(1).GeometryType = Inventor.Curve2dTypeEnum.kCircleCurve2d Then
            circleCurves.Add(c)
        End If
    End If
Next

Dim oDims As Inventor.GeneralDimensions = oSheet.DrawingDimensions.GeneralDimensions

' ===========================================
' Add overall horizontal dimension (length)
' ===========================================
If horizontalEdges.Count > 1 Then ' Need at least two edges for linear dimension
    Dim leftmostCurve As Inventor.DrawingCurve = Nothing
    Dim rightmostCurve As Inventor.DrawingCurve = Nothing
    Dim minX As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    
    For Each edge In horizontalEdges
        If TypeOf Edge.Segments.Item(1).Geometry Is Inventor.LineSegment2d Then
            Dim lineSeg As Inventor.LineSegment2d = Edge.Segments.Item(1).Geometry
            ' Consider both start and end points to find the true min/max X of the entire edge set
            If lineSeg.StartPoint.X < minX Then
                minX = lineSeg.StartPoint.X
                leftmostCurve = Edge ' Store the curve associated with this point
            End If
            If lineSeg.EndPoint.X < minX Then
                minX = lineSeg.EndPoint.X
                leftmostCurve = Edge ' Store the curve associated with this point
            End If

            If lineSeg.StartPoint.X > maxX Then
                maxX = lineSeg.StartPoint.X
                rightmostCurve = Edge ' Store the curve associated with this point
            End If
            If lineSeg.EndPoint.X > maxX Then
                maxX = lineSeg.EndPoint.X
                rightmostCurve = Edge ' Store the curve associated with this point
            End If
        End If
    Next

    If leftmostCurve IsNot Nothing AndAlso rightmostCurve IsNot Nothing Then
        ' Calculate average Y of the two extreme points for text placement
        ' This is simplified, assuming the horizontal edges are roughly at the same Y
        ' Ensure the point for dim text origin is created in the sheet context
        Dim dimPointH As Inventor.Point2d = tg.CreatePoint2d((minX + maxX) / 2, baseView.Position.Y - (baseView.Height / 2) - 2)
        
        Try
            ' Use the overload that takes two DrawingCurve objects directly
            oDims.AddLinear(leftmostCurve, rightmostCurve, dimPointH, Inventor.DimensionTypeEnum.kHorizontalDimensionType)
        Catch ex As Exception
            MessageBox.Show("Error adding horizontal dimension: " & ex.Message, "Dimension Error")
        End Try
    End If
End If

' ===========================================
' Add overall vertical dimension (width/height)
' ===========================================
If verticalEdges.Count > 1 Then
    Dim bottommostCurve As Inventor.DrawingCurve = Nothing
    Dim topmostCurve As Inventor.DrawingCurve = Nothing
    Dim minY As Double = Double.MaxValue
    Dim maxY As Double = Double.MinValue

    For Each edge In verticalEdges
        If TypeOf Edge.Segments.Item(1).Geometry Is Inventor.LineSegment2d Then
            Dim lineSeg As Inventor.LineSegment2d = Edge.Segments.Item(1).Geometry
            If lineSeg.StartPoint.Y < minY Then
                minY = lineSeg.StartPoint.Y
                bottommostCurve = Edge
            End If
            If lineSeg.EndPoint.Y < minY Then
                minY = lineSeg.EndPoint.Y
                bottommostCurve = Edge
            End If

            If lineSeg.StartPoint.Y > maxY Then
                maxY = lineSeg.StartPoint.Y
                topmostCurve = Edge
            End If
            If lineSeg.EndPoint.Y > maxY Then
                maxY = lineSeg.EndPoint.Y
                topmostCurve = Edge
            End If
        End If
    Next

    If bottommostCurve IsNot Nothing AndAlso topmostCurve IsNot Nothing Then
        ' This is simplified, assuming the vertical edges are roughly at the same X
        ' Ensure the point for dim text origin is created in the sheet context
        Dim dimPointV As Inventor.Point2d = tg.CreatePoint2d(baseView.Position.X - (baseView.Width / 2) - 2, (minY + maxY) / 2)
        
        Try
            oDims.AddLinear(bottommostCurve, topmostCurve, dimPointV, Inventor.DimensionTypeEnum.kVerticalDimensionType)
        Catch ex As Exception
            MessageBox.Show("Error adding vertical dimension: " & ex.Message, "Dimension Error")
        End Try
    End If
End If

' ===========================================
' Add thickness dimension (from the right view)
' ===========================================
' Thickness is represented by horizontal lines in the right view
Dim rightViewCurves As Inventor.DrawingCurvesEnumerator = rightView.DrawingCurves
Dim thicknessHorizontalEdges As New List(Of Inventor.DrawingCurve) 

For Each c As Inventor.DrawingCurve In rightViewCurves
    If c.Segments.Count > 0 AndAlso TypeOf c.Segments.Item(1).Geometry Is Inventor.LineSegment2d Then
        Dim lineSeg As Inventor.LineSegment2d = c.Segments.Item(1).Geometry
        If Math.Abs(lineSeg.StartPoint.Y - lineSeg.EndPoint.Y) < 0.01 Then ' Horizontal line in right view
            thicknessHorizontalEdges.Add(c)
        End If
    End If
Next

If thicknessHorizontalEdges.Count > 1 Then
    Dim thicknessLeftCurve As Inventor.DrawingCurve = Nothing
    Dim thicknessRightCurve As Inventor.DrawingCurve = Nothing
    Dim thicknessMinX As Double = Double.MaxValue
    Dim thicknessMaxX As Double = Double.MinValue

    For Each edge As Inventor.DrawingCurve In thicknessHorizontalEdges
        If TypeOf Edge.Segments.Item(1).Geometry Is Inventor.LineSegment2d Then
            Dim lineSeg As Inventor.LineSegment2d = Edge.Segments.Item(1).Geometry
            If lineSeg.StartPoint.X < thicknessMinX Then
                thicknessMinX = lineSeg.StartPoint.X
                thicknessLeftCurve = Edge
            End If
            If lineSeg.EndPoint.X < thicknessMinX Then
                thicknessMinX = lineSeg.EndPoint.X
                thicknessLeftCurve = Edge
            End If

            If lineSeg.StartPoint.X > thicknessMaxX Then
                thicknessMaxX = lineSeg.StartPoint.X
                thicknessRightCurve = Edge
            End If
            If lineSeg.EndPoint.X > thicknessMaxX Then
                thicknessMaxX = lineSeg.EndPoint.X
                thicknessRightCurve = Edge
            End If
        End If
    Next 

    If thicknessLeftCurve IsNot Nothing AndAlso thicknessRightCurve IsNot Nothing Then
        ' Simplified average Y for text placement
        Dim dimPointT As Inventor.Point2d = tg.CreatePoint2d(rightView.Position.X + (rightView.Width / 2) + 2, (thicknessLeftCurve.Segments.Item(1).Geometry.StartPoint.Y + thicknessRightCurve.Segments.Item(1).Geometry.StartPoint.Y) / 2)
        
        Try
            oDims.AddLinear(thicknessLeftCurve, thicknessRightCurve, dimPointT, Inventor.DimensionTypeEnum.kHorizontalDimensionType) ' Horizontal in the right view
        Catch ex As Exception
            MessageBox.Show("Error adding thickness dimension: " & ex.Message, "Dimension Error")
        End Try
    End If
End If 

' ===========================================
' Add chamfer dimensions
' ===========================================
For Each c As Inventor.DrawingCurve In chamferCurves
    If c.Segments.Count > 0 AndAlso TypeOf c.Segments.Item(1).Geometry Is Inventor.LineSegment2d Then ' Ensure it's a line segment
        Try
            ' Chamfer dimensions often take the DrawingCurve directly.
            Dim chamferPoint As Inventor.Point2d = tg.CreatePoint2d(c.Segments.Item(1).Geometry.StartPoint.X + 2, c.Segments.Item(1).Geometry.StartPoint.Y + 2)
            oSheet.DrawingDimensions.ChamferDimensions.Add(c, chamferPoint)
        Catch ex As Exception
            MessageBox.Show("Error adding chamfer dimension for curve " & c.ToString() & ": " & ex.Message, "Chamfer Dimension Error")
        End Try
    End If
Next

' ===========================================
' Add radial dimensions for holes
' ===========================================
For Each c As Inventor.DrawingCurve In circleCurves
    If c.Segments.Count > 0 AndAlso TypeOf c.Segments.Item(1).Geometry Is Inventor.Circle2d Then
        Try
            Dim circleGeom As Inventor.Circle2d = c.Segments.Item(1).Geometry
            Dim holePoint As Inventor.Point2d = tg.CreatePoint2d(circleGeom.Center.X + circleGeom.Radius + 1, circleGeom.Center.Y + circleGeom.Radius + 1)
            oSheet.DrawingDimensions.RadialDimensions.Add(c, holePoint)
        Catch ex As Exception
            MessageBox.Show("Error adding radial dimension for circle " & c.ToString() & ": " & ex.Message, "Radial Dimension Error")
        End Try
    End If
Next

' ===========================================
' Update Title Block
' ===========================================
Dim oTitleBlock As Inventor.TitleBlock = Nothing
Try
    If oSheet.TitleBlock IsNot Nothing Then
        oTitleBlock = oSheet.TitleBlock
        oTitleBlock.SetPromptResultText("TITLE", "PLATE AUTOMATION")
        oTitleBlock.SetPromptResultText("DRAWN BY", System.Environment.UserName)
        oTitleBlock.SetPromptResultText("DATE", Now.ToShortDateString)
    Else
        MessageBox.Show("No title block found on the sheet.", "Title Block Warning")
    End If
Catch ex As Exception
    MessageBox.Show("Error updating title block: " & ex.Message, "Title Block Error")
End Try

' ===========================================
' Export to PDF & DWG
' ===========================================
Dim path As String = System.IO.Path.GetDirectoryName(oDoc.FullFileName)
Dim nameOnly As String = System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)

Dim pdfName As String = System.IO.Path.Combine(path, nameOnly & ".pdf")
Dim pdfAddIn As Inventor.TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}") ' PDF Translator Add-In ID
Dim pdfContext As Inventor.TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext()
Dim pdfOptions As Inventor.NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
pdfContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism
If pdfAddIn.HasSaveCopyAsOptions(oDrawingDoc, pdfContext, pdfOptions) Then
    pdfOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
End If
Dim pdfData As Inventor.DataMedium = ThisApplication.TransientObjects.CreateDataMedium()
pdfData.FileName = pdfName
Try
    pdfAddIn.SaveCopyAs(oDrawingDoc, pdfContext, pdfOptions, pdfData)
Catch ex As Exception
    MessageBox.Show("Error exporting to PDF: " & ex.Message, "Export Error")
End Try


Dim dwgName As String = System.IO.Path.Combine(path, nameOnly & ".dwg")
Try
    oDrawingDoc.SaveAs(dwgName, False)
Catch ex As Exception
    MessageBox.Show("Error exporting to DWG: " & ex.Message, "Export Error")
End Try


ThisApplication.ActiveView.Fit
MessageBox.Show("Drawing created and exported to PDF and DWG at: " & path, "Success")

End Sub