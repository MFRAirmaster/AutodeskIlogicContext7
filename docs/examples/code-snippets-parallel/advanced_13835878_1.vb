' Title: CreateGeometryIntent creates kNoPointIntent type
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/creategeometryintent-creates-knopointintent-type/td-p/13835878
' Category: advanced
' Scraped: 2025-10-07T13:58:09.669013

Public Function IncludeWorkFeatureAndGetIntent(view As DrawingView,
                                                   partName As String,
                                                   featureName As String,
                                          Optional include As Boolean = True,
                                          Optional containsMatch As Boolean = True,
                                          Optional searchAxes As Boolean = False,
                                          Optional searchPlanes As Boolean = False,
                                          Optional searchPoints As Boolean = False,
                                          Optional search2DSketches As Boolean = False,
                                          Optional search3DSketches As Boolean = False,
                                          Optional searchModelAnnotations As Boolean = False,
                                          Optional caseInsensitive As Boolean = True,
                                          Optional onlyFromVisibleOccurrences As Boolean = False
                ) As GeometryIntent

        If view Is Nothing OrElse view.ReferencedDocumentDescriptor Is Nothing Then Return Nothing

        Dim refDoc As Document = view.ReferencedDocumentDescriptor.ReferencedDocument
        Dim sheet As Sheet = view.Parent
        Dim cmp = If(caseInsensitive, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)

        Dim matchStr As Func(Of String, String, Boolean) = Function(src As String, pat As String) As Boolean
                                                               If String.IsNullOrEmpty(pat) Then Return True
                                                               If containsMatch Then
                                                                   Return src?.IndexOf(pat, cmp) >= 0
                                                               Else
                                                                   Return String.Equals(src, pat, cmp)
                                                               End If
                                                           End Function

        Dim baseOf As Func(Of String, String) = Function(nm As String) As String
                                                    If String.IsNullOrEmpty(nm) Then Return nm
                                                    Dim p = nm.IndexOf(":"c)
                                                    Return If(p >= 0, nm.Substring(0, p), nm)
                                                End Function

        ' ---- PART VIEW ----
        Dim prt = TryCast(refDoc, PartDocument)
        If prt IsNot Nothing Then
            If Not matchStr(IO.Path.GetFileNameWithoutExtension(prt.DisplayName), partName) AndAlso
           Not matchStr(prt.DisplayName, partName) Then Return Nothing

            Dim pDef = prt.ComponentDefinition

            If searchAxes Then
                For Each ax In pDef.WorkAxes
                    If matchStr(ax.Name, featureName) Then
                        Try
                            view.SetIncludeStatus(ax, include)
                            Return sheet.CreateGeometryIntent(ax)
                        Catch : End Try
                    End If
                Next
            End If

            If searchPlanes Then
                For Each pl In pDef.WorkPlanes
                    If matchStr(pl.Name, featureName) Then
                        Try
                            view.SetIncludeStatus(pl, include)
                            Return sheet.CreateGeometryIntent(pl)
                        Catch
                        End Try
                    End If
                Next
            End If

            If searchPoints Then
                For Each wp In pDef.WorkPoints
                    If matchStr(wp.Name, featureName) Then
                        Try
                            view.SetIncludeStatus(wp, include)
                            Return sheet.CreateGeometryIntent(wp)
                        Catch
                        End Try
                    End If
                Next
            End If

            If search2DSketches Then
                For Each sk In pDef.Sketches
                    If matchStr(sk.Name, featureName) Then
                        Try
                            view.SetIncludeStatus(sk, include)
                            Return sheet.CreateGeometryIntent(sk)
                        Catch
                        End Try
                    End If
                Next
            End If

            If search3DSketches Then
                For Each sk3 In pDef.Sketches3D
                    If matchStr(sk3.Name, featureName) Then
                        Try
                            view.SetIncludeStatus(sk3, include)
                            Return sheet.CreateGeometryIntent(sk3)
                        Catch
                        End Try
                    End If
                Next
            End If

            Return Nothing
        End If
        ' ---- ASSEMBLY VIEW ----
        Dim asm = TryCast(refDoc, AssemblyDocument)
        If asm Is Nothing Then Return Nothing

        Dim occsEnum As ComponentOccurrencesEnumerator = asm.ComponentDefinition.Occurrences.AllReferencedOccurrences(asm.ComponentDefinition)

        For Each occ As ComponentOccurrence In occsEnum
            If occ Is Nothing OrElse occ.Suppressed Then Continue For

            Dim occDoc As Document = TryCast(occ.Definition.Document, Document)
            Dim occDocBase As String = ""
            If occDoc IsNot Nothing Then
                occDocBase = IO.Path.GetFileNameWithoutExtension(occDoc.DisplayName)
            End If
            If Not matchStr(occDocBase, partName) Then Continue For

            If onlyFromVisibleOccurrences Then
                Dim hasCurves As Boolean = False
                Try
                    hasCurves = (view.DrawingCurves(occ)?.Count > 0)
                Catch
                    hasCurves = False
                End Try
                If Not hasCurves Then
                    Continue For
                End If
            End If
            Dim pDef = Nothing
            pDef = TryCast(occ.Definition, PartComponentDefinition)
            If pDef Is Nothing Then
                pDef = TryCast(occ.Definition, AssemblyComponentDefinition)
            End If
            If pDef Is Nothing Then Continue For

            ' --- AXES ---
            If searchAxes Then
                For Each ax In pDef.WorkAxes
                    If Not matchStr(ax.Name, featureName) Then Continue For
                    Dim px As WorkAxisProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(ax, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch
                    End Try
                Next
            End If

            ' --- PLANES ---
            If searchPlanes Then
                For Each pl In pDef.WorkPlanes
                    If Not matchStr(pl.Name, featureName) Then Continue For
                    Dim px As WorkPlaneProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(pl, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch
                    End Try
                Next
            End If

            ' --- POINTS ---
            If searchPoints Then
                For Each wp In pDef.WorkPoints
                    If Not matchStr(wp.Name, featureName) Then Continue For
                    Dim px As WorkPointProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(wp, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch
                    End Try
                Next
            End If

            ' --- 2D SKETCHES ---
            If search2DSketches Then
                For Each sk As PlanarSketch In pDef.Sketches
                    If Not matchStr(sk.Name, featureName) Then Continue For
                    Dim px As PlanarSketchProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(sk, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch
                    End Try
                Next
            End If
            ' --- 3D SKETCHES ---
            If search3DSketches Then
                For Each sk3 In pDef.Sketches3D
                    If Not matchStr(sk3.Name, featureName) Then Continue For
                    Dim px As Sketch3DProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(sk3, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch
                    End Try
                Next
            End If

        Next

        Return Nothing
    End Function