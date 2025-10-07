' Title: Colour Coding Parts by Length in Drawings
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/colour-coding-parts-by-length-in-drawings/td-p/13802556
' Category: advanced
' Scraped: 2025-10-07T14:10:10.634083

Sub main
	On Error Resume Next
	Dim view_repr_name = "color_coding"
	Dim  custom_iproperty_name As String ="s_length"
	Dim color_dictionary As New Dictionary(Of String, String) From {
 {"1000 mm",  "xYellow;255;255;0"},
 {"1100 mm",  "xRed;255;0;0"},
 {"1200 mm", "xBlue;0;0;255" }}
 If ThisApplication.ActiveDocumentType =DocumentTypeEnum.kPartDocumentObject Then
	Dim pdoc As PartDocument = ThisDoc.Document
	If color_dictionary.ContainsKey(pdoc.PropertySets(4)(custom_iproperty_name).Value) Then
		Dim name As String=""
		Dim  R, G, B  As Integer
		name = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value), ";")(0)
		R = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value), ";")(1)
		G = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value), ";")(2)
		B = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value),";")(3)
		Call GetColorAsset(pdoc,view_repr_name, name, R, G, B)
	End If
ElseIf ThisApplication.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
	Dim ass As AssemblyDocument =ThisApplication.ActiveDocument
	Dim check As String = ""
		check=ass.ComponentDefinition.RepresentationsManager.DesignViewRepresentations(view_repr_name).Name 
		If check = view_repr_name Then
			ass.ComponentDefinition.RepresentationsManager.DesignViewRepresentations(view_repr_name).Activate 
		Else
			ass.ComponentDefinition.RepresentationsManager.DesignViewRepresentations.Add(view_repr_name).Activate 
		End If
		For Each occ As ComponentOccurrence In ass.ComponentDefinition.Occurrences
			If occ.DefinitionDocumentType<>DocumentTypeEnum.kPartDocumentObject Then Continue For
			Dim pdoc As PartDocument = occ.Definition.Document
			If color_dictionary.ContainsKey(pdoc.PropertySets(4)(custom_iproperty_name).Value) Then
				Dim name As String
				Dim  R, G, B  As Integer
				name = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value), ";")(0)
				R = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value), ";")(1)
				G = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value), ";")(2)
				B = Split(color_dictionary(pdoc.PropertySets(4)(custom_iproperty_name).Value),";")(3)
				Call GetColorAsset(pdoc,view_repr_name, name, R, G, B)
				occ.SetDesignViewRepresentation(view_repr_name,, True)
			End If
			
		Next
End If

End Sub
 Function GetColorAsset(doc As PartDocument,view_repr_name As String, name As String, R As Integer, G As Integer, B As Integer) As Asset
Dim Asset As Asset=Nothing
	Try
		Asset = doc.Assets.Add(AssetTypeEnum.kAssetTypeAppearance, "Generic", , name)

	Catch ex As Exception
		Try
			Asset = doc.Assets.Item(name)
		Catch ex1 As Exception
			doc.Rebuild()
			doc.Update()
			Try
				Asset = doc.Assets.Add(AssetTypeEnum.kAssetTypeAppearance, "Generic", , name)
			Catch ex2 As Exception
			End Try
		End Try
	End Try
Dim uColor As ColorAssetValue = Asset.Item("generic_diffuse")
Dim clr As Inventor.Color = ThisApplication.TransientObjects.CreateColor(R, G, B)
uColor.Value = clr
check  = ""
Try
check = doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations(view_repr_name).Name 
Catch ex As Exception
End Try
If check = view_repr_name Then
	doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations(view_repr_name).Activate 
Else
	doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations.Add(view_repr_name).Activate 
End If
Try
doc.ActiveAppearance = Asset
Catch ex As Exception
End Try
doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations(1).Activate
doc.Save		
End Function