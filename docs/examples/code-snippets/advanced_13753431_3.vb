' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431
' Category: advanced
' Scraped: 2025-10-07T13:07:50.456065

Sub Main()
	
	Dim oDoc As AssemblyDocument = ThisDoc.Document
	
	'set up face collection 
	Dim oFace As Face
	Dim oFaceCol As FaceCollection = ThisApplication.TransientObjects.CreateFaceCollection
	
	For Each oComponent In oDoc.ComponentDefinition.Occurrences
		
		If oComponent.Suppressed = True Then Continue For
			
		If TypeOf oComponent.Definition Is VirtualComponentDefinition Then Continue For
		
		GetFaces(oComponent, oFaceCol)
		
	Next
	
	For Each oFace In oFaceCol
		
		ThisApplication.CommandManager.DoSelect(oFace)
		
	Next
	
	MsgBox("Number of faces: " & oFaceCol.Count)
	
	'Create finish definition.
	Dim oFinish As FinishFeature
	Dim oFinishColour = iProperties.Value("Custom", "Box Colour")
	Dim oAppearance As Object = oDoc.AppearanceAssets.Item(oFinishColour)
	Dim oFinishFeatures As FinishFeatures = oDoc.ComponentDefinition.Features.FinishFeatures
    Dim oFinishDef As FinishDefinition = oFinishFeatures.CreateFinishDefinition(oFaceCol, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)
	
    'Create finish feature and rename
	Try
	    oFinish = oFinishFeatures.Add(oFinishDef)
		'oFinish.Name = oFinishName
	Catch
		Logger.Trace("Couldn't create finish in " & oDoc.DisplayName & ".")
	End Try
	
End Sub

Sub GetFaces(oOcc As ComponentOccurrence, oFaceCollection As FaceCollection)
		
		'check if occurance is assembly
		If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
			
			For Each oComponent In oOcc.SubOccurrences
				
				'skip if component component is suppressed
				If oComponent.Suppressed = True Then Continue For
				
				'skip if virtual component
				If TypeOf oComponent.Definition Is VirtualComponentDefinition Then Continue For
				
				GetFaces(oComponent, oFaceCollection) 
				
			Next
			
		Else
		
		    For Each oSurfaceBody In oOcc.SurfaceBodies
		        For Each oFace In oSurfaceBody.Faces
		            oFaceCollection.Add(oFace)
		        Next
		    Next
			
		End If
	
End Sub