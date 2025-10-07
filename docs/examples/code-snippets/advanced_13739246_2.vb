' Title: Balloon all components of an assembly view
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/balloon-all-components-of-an-assembly-view/td-p/13739246
' Category: advanced
' Scraped: 2025-10-07T13:29:54.210121

' Add a balloon for each referenced component
Dim oInvApp As Inventor.Application = ThisApplication
Dim oDrawDoc As DrawingDocument = oInvApp.ActiveDocument
Dim oSheet As Sheet = oDrawDoc.ActiveSheet
Dim oView As DrawingView
Dim oTG As TransientGeometry = oInvApp.TransientGeometry
Dim oTO As TransientObjects = oInvApp.TransientObjects

' Get the target view
For Each v As DrawingView In oSheet.DrawingViews
    If v.Name = "TOP" Then
        oView = v
        Exit For
    End If
Next
If oView Is Nothing Then Exit Sub

Dim oAssemblyDoc As AssemblyDocument = oView.ReferencedDocumentDescriptor.ReferencedDocument
Dim oOccs As Inventor.ComponentOccurrences = oAssemblyDoc.ComponentDefinition.Occurrences

For Each oOcc As ComponentOccurrence In oOccs
	Dim oLeaderPoints As ObjectCollection = oTO.CreateObjectCollection
    Try
        ' Try each edge until you find a drawable curve
        For Each oEdge As Edge In oOcc.SurfaceBodies.Item(1).Edges
            Dim oEdgeProxy As EdgeProxy
            oOcc.CreateGeometryProxy(oEdge, oEdgeProxy)

            Dim oCurves As DrawingCurvesEnumerator = oView.DrawingCurves(oEdgeProxy)
            If oCurves.Count > 0 Then
                Dim oCurve As DrawingCurve = oCurves.Item(1)
                Dim midPt As Point2d = oCurve.MidPoint
                oLeaderPoints.Add(oTG.CreatePoint2d(midPt.X, midPt.Y))
                Dim intent As GeometryIntent = oSheet.CreateGeometryIntent(oCurve, 0.5)
                oLeaderPoints.Add(intent)
				oSheet.Balloons.Add(oLeaderPoints)
				Logger.Info("Point added for: " & oOcc.Name)
                Exit For
            End If
        Next
    Catch ex As Exception
        Logger.Error("Curve not found for " & oOcc.Name & ": " & ex.Message)
    End Try
Next