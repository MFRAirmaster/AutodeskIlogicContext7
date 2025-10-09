' Title: iLogic Code to Delete Specific Parameter in all Parts Within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-delete-specific-parameter-in-all-parts-within/td-p/13803275#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:55:10.667585

Sub Main

	Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
	Dim oOccs As ComponentOccurrences = oADoc.ComponentDefinition.Occurrences

	Call TraverseAssembly(oOccs)

End Sub

Function TraverseAssembly(oOccs As ComponentOccurrences)

	Dim oOcc As ComponentOccurrence
	For Each oOcc In oOccs
		If oOcc.Suppressed = True Then Continue For
		'get the document from the occurrence
		Dim oDoc = oOcc.Definition.Document

		Try
			oDoc.ComponentDefinition.Parameters.Item("Thickness").Delete
			Logger.Info(oOcc.Name & " parameter deleted")
		Catch 
			Logger.Info(oOcc.Name & " could not find or delete parameter")
		End Try

		'if assembly loop back up and step down into it's occurrences
		If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
			Call TraverseAssembly(oOcc.SubOccurrences)
		Else 
			' handle stuff for just parts here if needed
		End If
	Next

End Function