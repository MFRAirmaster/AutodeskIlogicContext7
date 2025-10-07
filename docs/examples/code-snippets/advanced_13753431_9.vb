' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431
' Category: advanced
' Scraped: 2025-10-07T13:07:50.456065

Sub Main
	
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oADoc As AssemblyDocument = TryCast(ThisDoc.FactoryDocument, Inventor.AssemblyDocument)
	If oADoc Is Nothing Then Return
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition

	'set list of components to skip
	Dim oSkipList As New ArrayList
	oSkipList.Clear 
	oSkipList.Add("20110")
	oSkipList.Add("20224")
	
	Dim oFaceColl As Inventor.FaceCollection
	
	For Each oComponent In oADoc.ComponentDefinition.Occurrences
		
		If oComponent.Suppressed = True Then Continue For
		If TypeOf oComponent.Definition Is VirtualComponentDefinition Then Continue For
		
		GetFaces(oComponent, oFaceColl, oSkipList)
		
	Next
	
	Dim oFinishColour = iProperties.Value("Custom", "Box Colour")
	Dim oAppearance As Object = oADoc.AppearanceAssets.Item(oFinishColour)
	
	Dim oFinishFeats As FinishFeatures = oADef.Features.FinishFeatures
	Dim oFinishDef As FinishDefinition = oFinishFeats.CreateFinishDefinition(oFaceColl, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)
	Dim oFinishFeat As FinishFeature = oFinishFeats.Add(oFinishDef)
	
End Sub


Function GetFaces(occ As Inventor.ComponentOccurrence, Optional ByRef faceColl As Inventor.FaceCollection = Nothing, Optional oSkipList As ArrayList = Nothing) As Inventor.FaceCollection
	
	If (occ Is Nothing) OrElse occ.Suppressed Then Return Nothing
	If TypeOf occ.Definition Is VirtualComponentDefinition Then Return Nothing
	
	Dim oOccName As String = Split(occ.Name, ":")(0)
	If oSkipList.Contains(oOccName) Then Return Nothing
	
	If faceColl Is Nothing Then faceColl = ThisApplication.TransientObjects.CreateFaceCollection()
		
	For Each oBody As SurfaceBody In occ.SurfaceBodies
		For Each oFace As Face In oBody.Faces
			faceColl.Add(oFace)
		Next 'oFace
	Next 'oBody
	
	If occ.SubOccurrences IsNot Nothing AndAlso occ.SubOccurrences.Count > 0 Then
		For Each oSubOcc As Inventor.ComponentOccurrence In occ.SubOccurrences
			GetFaces(oSubOcc, faceColl, oSkipList)
		Next 'oSubOcc
	End If
	
	Return faceColl
	
End Function