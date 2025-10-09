' Title: Model States and BOMstructure
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/model-states-and-bomstructure/td-p/13772152
' Category: advanced
' Scraped: 2025-10-09T09:03:14.421530

Dim oAssyDoc As AssemblyDocument = ThisDoc.Document
For Each oRefDoc As Document In oAssyDoc.AllReferencedDocuments
	If oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then ' = .iam
		If oRefDoc.ComponentDefinition.ModelStates.Count > 1 Then
			oRefDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPhantomBOMStructure
		End If
	End If
Next