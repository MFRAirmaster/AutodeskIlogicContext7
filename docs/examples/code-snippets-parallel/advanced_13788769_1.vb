' Title: About the issue of obtaining the extrusion start face
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/about-the-issue-of-obtaining-the-extrusion-start-face/td-p/13788769
' Category: advanced
' Scraped: 2025-10-07T14:01:45.468718

Private Function HandleDistanceFromFaceExtent(oExtrude As ExtrudeFeature) As Boolean
        Try
            Dim oDistFace As DistanceFromFaceExtent = CType(oExtrude.Extent, DistanceFromFaceExtent)
            Dim dirs As PartFeatureExtentDirectionEnum() = {
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
                PartFeatureExtentDirectionEnum.kSymmetricExtentDirection
            }

            Dim currentDir As PartFeatureExtentDirectionEnum = oDistFace.Direction
            Dim idx As Integer = Array.IndexOf(dirs, currentDir)
            Dim newDir As PartFeatureExtentDirectionEnum = dirs((If(idx = -1, 0, idx + 1) Mod dirs.Length))

            Dim fromFace As Object = oDistFace.FromFace
            Dim distance As Object = oDistFace.Distance.Expression

            oExtrude.Definition.SetDistanceFromFaceExtent(fromFace, True, newDir, distance)
            Return True
        Catch
            Return False
        End Try
    End Function