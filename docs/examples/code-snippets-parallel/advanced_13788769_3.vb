' Title: About the issue of obtaining the extrusion start face
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/about-the-issue-of-obtaining-the-extrusion-start-face/td-p/13788769
' Category: advanced
' Scraped: 2025-10-09T09:08:55.479773

Public Sub Main()
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oPDoc As PartDocument = TryCast(oInvApp.ActiveDocument, PartDocument)
	If HandleDistanceFromFaceExtent(oPDoc.ComponentDefinition.Features(1)) Then
		Logger.Info("GOOD!")
	Else
		Logger.Info("NOT GOOD!")
	End If	
End Sub

Private Function HandleDistanceFromFaceExtent(oExtrude As ExtrudeFeature) As Boolean
    Try
        Dim oDistFace As PartFeatureExtent = CType(oExtrude.Definition.Extent, PartFeatureExtent)
        Dim dirs As PartFeatureExtentDirectionEnum() = {
            PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            PartFeatureExtentDirectionEnum.kSymmetricExtentDirection
        }

        Dim currentDir As PartFeatureExtentDirectionEnum = oDistFace.Direction
        Dim idx As Integer = Array.IndexOf(dirs, currentDir)
        Dim newDir As PartFeatureExtentDirectionEnum = dirs((If (idx = -1, 0, idx + 1) Mod dirs.Length))
		
	    Dim distance As Object = oDistFace.Distance.Expression
		If oDistFace.Type = 83917824 Then
			oExtrude.Definition.SetDistanceExtent(distance, newDir)
		Else
	        Dim fromFace As Object = oDistFace.FromFace	
	        oExtrude.Definition.SetDistanceFromFaceExtent(fromFace, True, newDir, distance)
		End If
        Return True
    Catch
        Return False
    End Try
End Function