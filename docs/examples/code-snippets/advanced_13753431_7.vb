' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431
' Category: advanced
' Scraped: 2025-10-07T13:07:50.456065

Function GetFaces(occ As Inventor.ComponentOccurrence, _
	Optional ByRef faceColl As Inventor.FaceCollection = Nothing) As Inventor.FaceCollection
	If (occ Is Nothing) OrElse occ.Suppressed Then Return Nothing
	If TypeOf occ.Definition Is VirtualComponentDefinition Then Return Nothing
	If faceColl Is Nothing Then faceColl = ThisApplication.TransientObjects.CreateFaceCollection()
	For Each oBody As SurfaceBody In occ.SurfaceBodies
		For Each oFace As Face In oBody.Faces
			faceColl.Add(oFace)
		Next 'oFace
	Next 'oBody
	If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
		For Each oSubOcc As Inventor.ComponentOccurrence In occ.SubOccurrences
			GetFaces(oSubOcc, faceColl)
		Next 'oSubOcc
	End If
	Return faceColl
End Function