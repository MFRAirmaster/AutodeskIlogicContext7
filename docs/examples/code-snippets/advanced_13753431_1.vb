' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431
' Category: advanced
' Scraped: 2025-10-07T13:07:50.456065

Sub Main()
	
	Dim oAsm As AssemblyDocument = ThisDoc.Document
	Dim oFaceCol As FaceCollection = ThisApplication.TransientObjects.CreateFaceCollection
	Dim oFaceProx As FaceProxy
	Dim oComponent As ComponentOccurrence = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "select subassembly")
	
	If oComponent.Definition.Type = ObjectTypeEnum.kAssemblyComponentDefinitionObject
		Dim oDef As AssemblyComponentDefinition = oComponent.Definition
		For Each oFaceProx In GetProxies(oComponent)
			ThisApplication.CommandManager.DoSelect(oFaceProx)
			oFaceCol.Add(oFaceProx)
		Next
	End If
	
	MsgBox(oFaceCol.Count)
	
	'Create finish definition.
	Dim oFinishColour = iProperties.Value("Custom", "Box Colour")
	Dim oAppearance As Object = oAsm.AppearanceAssets.Item(oFinishColour)
	Dim oFinishFeatures As FinishFeatures = oAsm.ComponentDefinition.Features.FinishFeatures
	Dim oFinishDef As FinishDefinition = oFinishFeatures.CreateFinishDefinition(oFaceCol, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)

'Create finish feature 
	Try
	    Dim oFinish As FinishFeature = oFinishFeatures.Add(oFinishDef)
	Catch
		Logger.Trace("Couldn't create finish in " & oAsm.DisplayName & ".")
	End Try

End Sub

Function GetProxies(oComponent As ComponentOccurrence) As ObjectCollection
	
	Dim oProxCol As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
	Dim oDef As AssemblyComponentDefinition = oComponent.Definition
	
	For Each oComp As ComponentOccurrence In oDef.Occurrences
		If oComp.Definition.Type = ObjectTypeEnum.kAssemblyComponentDefinitionObject
			For Each oProx1 As FaceProxy In GetProxies(oComp)
				Dim oProx2 As FaceProxy
				Call oComponent.CreateGeometryProxy(oProx1, oProx2)
				oProxCol.Add(oProx2)
			Next
		Else
			Dim oCompDef As PartComponentDefinition = oComp.Definition
			For Each oBod As SurfaceBody In oCompDef.SurfaceBodies
				For Each oFace As Face In oBod.Faces
				Dim oProx1 As FaceProxy
				Dim oProx2 As FaceProxy
				Call oComp.CreateGeometryProxy(oFace, oProx1)
				Call oComponent.CreateGeometryProxy(oProx1, oProx2)
				oProxCol.Add(oProx2)
				Next
			Next
		End If
	Next
	
	Return oProxCol
	
End Function