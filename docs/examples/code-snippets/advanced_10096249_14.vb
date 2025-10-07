' Title: Auto ballooning a drawing - Attach balloon to DrawingCurveSegment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-ballooning-a-drawing-attach-balloon-to-drawingcurvesegment/td-p/10096249#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:31:35.305789

Public Sub AddBalloonsForOccurrenceNames(view As DrawingView,
                                         partList As PartsList,
                                         targetNames As IEnumerable(Of String),
                                         Optional topOffsetCm As Double = 1.0)

        If view Is Nothing OrElse partList Is Nothing Then Exit Sub
        If targetNames Is Nothing Then Exit Sub

        Dim wanted As New HashSet(Of String)(targetNames, StringComparer.OrdinalIgnoreCase)
        Dim tg = g_inventorApplication.TransientGeometry
        Dim created As New List(Of Balloon)

        InitialiseViewBoundingBox(view)

        For Each row As PartsListRow In partList.PartsListRows
            If row Is Nothing OrElse row.Ballooned OrElse Not row.Visible Then Continue For

            ' ---- inlined GetBalloonAttachGeometryFiltered ----
            Dim gi As GeometryIntent = Nothing

            If wanted.Count > 0 Then
                ' All occurrences for this row's referenced file
                Dim refDocDesc As DocumentDescriptor = row.ReferencedFiles.Item(1).DocumentDescriptor
                Dim occAll As ComponentOccurrencesEnumerator =
                view.ReferencedDocumentDescriptor.ReferencedDocument.
                    ComponentDefinition.Occurrences.AllReferencedOccurrences(refDocDesc)

                ' Filter occurrences by base name in "wanted"
                Dim occFiltered As New List(Of ComponentOccurrence)
                For Each occ As ComponentOccurrence In occAll
                    If occ Is Nothing OrElse occ.Suppressed Then Continue For

                    ' ---- inlined NormalizeOccName ----
                    Dim nm As String = occ.Name
                    Dim baseName As String
                    If String.IsNullOrEmpty(nm) Then
                        baseName = nm
                    Else
                        Dim p As Integer = nm.IndexOf(":"c)
                        baseName = If(p >= 0, nm.Substring(0, p), nm)
                    End If
                    If wanted.Contains(baseName) Then occFiltered.Add(occ)
                Next

                If occFiltered.Count > 0 Then
                    ' ---- inlined GetCurvesFromOccasList ----
                    Dim curves As New List(Of DrawingCurve)
                    For Each o As ComponentOccurrence In occFiltered
                        If o Is Nothing OrElse o.Suppressed Then Continue For
                        For Each dc As DrawingCurve In view.DrawingCurves(o)
                            If Not curves.Contains(dc) Then curves.Add(dc)
                        Next
                    Next
                    If curves.Count > 0 Then
                        Dim seg As DrawingCurveSegment = GetBestSegmentFromOccurrence(curves)
                        If seg IsNot Nothing Then
                            Dim sheet As Sheet = seg.Parent.Parent.Parent
                            gi = sheet.CreateGeometryIntent(seg.Parent, GetSegmentMidPoint(seg))
                        End If
                    End If
                End If
            End If
            If gi Is Nothing Then Continue For

            Dim topY As Double = view.Top + topOffsetCm
            Dim midX As Double = view.Left + (view.Width / 2.0)

            Dim leaders As ObjectCollection = g_inventorApplication.TransientObjects.CreateObjectCollection
            leaders.Add(tg.CreatePoint2d(midX, topY))
            leaders.Add(gi)

            Dim b As Balloon = view.Parent.Balloons.Add(leaders)
            created.Add(b)
        Next

        If created.Count > 0 Then
            DistributeBalloonsAlongTopOrBottom(
            created,
            yPos:=view.Top + topOffsetCm,
            xStart:=view.Left,
            xEnd:=view.Left + view.Width)
        End If
    End Sub