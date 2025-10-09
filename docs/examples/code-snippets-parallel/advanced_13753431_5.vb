' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:09:48.375498

Sub Main
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oADoc As AssemblyDocument = TryCast(ThisDoc.FactoryDocument, Inventor.AssemblyDocument)
	If oADoc Is Nothing Then Return
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
	
	Dim oPickedOcc As ComponentOccurrence = oInvApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select a Component.")
	If oPickedOcc Is Nothing Then Return
	
	Dim oFaceColl1 As Inventor.FaceCollection = oInvApp.TransientObjects.CreateFaceCollection()
	GetFaces2(oPickedOcc, oFaceColl1)
	
'	Dim oFaceColl2 As Inventor.FaceCollection = GetFaces(oPickedOcc)
	
	Dim oAppAsset As Inventor.Asset = oADoc.AppearanceAssets.Item("Red")
	
	Dim oFinishFeats As FinishFeatures = oADef.Features.FinishFeatures
	Dim oFinishDef As FinishDefinition = oFinishFeats.CreateFinishDefinition( _
	oFaceColl1, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppAsset)
	Dim oFinishFeat As FinishFeature = oFinishFeats.Add(oFinishDef)
End Sub

'works OK when the component is a part, not a sub assembly
Function GetFaces(occ As Inventor.ComponentOccurrence) As Inventor.FaceCollection
	If (occ Is Nothing) OrElse occ.Suppressed Then Return Nothing
	If TypeOf occ.Definition Is VirtualComponentDefinition Then Return Nothing
	Dim oFaceColl As Inventor.FaceCollection = ThisApplication.TransientObjects.CreateFaceCollection()
	For Each oBody As SurfaceBody In occ.SurfaceBodies
		For Each oFace As Face In oBody.Faces
			oFaceColl.Add(oFace)
		Next 'oFace
	Next 'oBody
	Return oFaceColl
End Function

'works when the component is either a part, or sub assembly
Sub GetFaces2(occ As Inventor.ComponentOccurrence, ByRef faceColl As Inventor.FaceCollection)
	If (occ Is Nothing) OrElse occ.Suppressed Then Return
	If TypeOf occ.Definition Is VirtualComponentDefinition Then Return
	If faceColl Is Nothing Then faceColl = ThisApplication.TransientObjects.CreateFaceCollection()
	For Each oBody As SurfaceBody In occ.SurfaceBodies
		For Each oFace As Face In oBody.Faces
			faceColl.Add(oFace)
		Next 'oFace
	Next 'oBody
	If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
		For Each oSubOcc As Inventor.ComponentOccurrence In occ.SubOccurrences
			GetFaces2(oSubOcc, faceColl)
		Next
	End If
End Sub