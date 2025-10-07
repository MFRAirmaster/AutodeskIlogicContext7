' Title: Model States and BOMstructure
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/model-states-and-bomstructure/td-p/13772152#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:10:32.262577

'make sure our rule is working with 'FactoryDocument' if it has multiple ModelStates
Dim oAssyDoc As AssemblyDocument = TryCast(ThisDoc.FactoryDocument, AssemblyDocument)
'if current document was not an Assembly, it will be Nothing
If oAssyDoc Is Nothing Then Return
For Each oRefDoc As Document In oAssyDoc.AllReferencedDocuments
	If oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then ' = .iam
		'creating the AssemblyDocument variable enables Intellisense recognition of later properties
		Dim oRefADoc As AssemblyDocument = oRefDoc
		'might need to check or record name of controlling ModelState (if any)
		Dim sMSName As String = oRefADoc.ModelStateName
		'get 'Factory' version, if there is one (only version we can edit)
		If oRefADoc.ComponentDefinition.IsModelStateMember Then
			oRefADoc = oRefADoc.ComponentDefinition.FactoryDocument
		End If
		'update the referenced document, if it is needed (sometimes necessary)
		If oRefADoc.RequiresUpdate Then oRefADoc.Update2(True)
		'get its ModelStates (if it has any - none in iPart/iAssembly)
		Dim oMSs As ModelStates = oRefADoc.ComponentDefinition.ModelStates
		If oMSs.Count > 1 Then
			'may want to set MemberEditScope, to control area of effect
			oMSs.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
			'oMSs.MemberEditScope = MemberEditScopeEnum.kEditActiveMember
			oRefADoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPhantomBOMStructure
			'oRefADoc.Update2(True) this may not be necessary
		End If
	End If
Next 'oRefDoc
oAssyDoc.Update2(True)