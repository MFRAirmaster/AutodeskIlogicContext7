' Title: Help with iLogic: Projecting Bend Lines from Unfolded Sheet Metal
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/help-with-ilogic-projecting-bend-lines-from-unfolded-sheet-metal/td-p/13736222#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:09:52.033300

Sub Main()

    ' Setup
    Dim oDoc As PartDocument = ThisApplication.ActiveDocument
    Dim oCompDef As SheetMetalComponentDefinition = oDoc.ComponentDefinition
    Dim oTG As TransientGeometry = ThisApplication.TransientGeometry

    ' === STEP 1: Find the top unfolded flat face ===
    Dim topFace As Face = Nothing
    Dim maxArea As Double = 0

    For Each oFace As Face In oCompDef.SurfaceBodies(1).Faces
        If oFace.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
            Dim area As Double = oFace.Evaluator.Area
            If area > maxArea Then
                maxArea = area
                topFace = oFace
            End If
        End If
    Next

    If topFace Is Nothing Then
        MessageBox.Show("No flat top face found. Please unfold the part first.")
        Return
    End If

    ' === STEP 2: Create sketch on top face ===
    Dim oSketch As PlanarSketch = oCompDef.Sketches.Add(topFace)

    ' === STEP 3: Loop through edges and detect true bend lines ===
    For Each oEdge As Edge In oCompDef.SurfaceBodies(1).Edges

        If oEdge.GeometryType = CurveTypeEnum.kLineSegmentCurve Then

            If oEdge.Faces.Count = 2 Then

                Dim face1 As Face = oEdge.Faces(1)
                Dim face2 As Face = oEdge.Faces(2)

                ' Both must be planar
                If face1.SurfaceType = SurfaceTypeEnum.kPlaneSurface And face2.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then

                    ' Exclude edge if it lies between top face and another (outer edge)
                    If face1.Equals(topFace) Or face2.Equals(topFace) Then
                        Continue For
                    End If

                    ' Check angle between faces
                    Dim normal1 As UnitVector = GetFaceNormal(face1)
                    Dim normal2 As UnitVector = GetFaceNormal(face2)

                    Dim dot As Double = normal1.DotProduct(normal2)

                    If Math.Abs(dot) < 0.98 Then
                        ' True bend line — project it
                        oSketch.AddByProjectingEntity(oEdge)
                    End If

                End If
            End If
        End If
    Next

    oSketch.Name = "TrueBendLines"
    MessageBox.Show("Bend lines (excluding outer edges) projected to sketch.")

End Sub

' === Helper: Safe face normal evaluator ===
Function GetFaceNormal(oFace As Face) As UnitVector
    Dim pt As Point = oFace.PointOnFace
    Dim oPointArr(2) As Double
    oPointArr(0) = pt.X
    oPointArr(1) = pt.Y
    oPointArr(2) = pt.Z

    Dim oNormalArr(2) As Double
    oFace.Evaluator.GetNormalAtPoint(oPointArr, oNormalArr)

    Dim oVec As Vector = ThisApplication.TransientGeometry.CreateVector(oNormalArr(0), oNormalArr(1), oNormalArr(2))
    Return oVec.AsUnitVector
End Function