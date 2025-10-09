' Title: Utilizing MakePath in SelectSet.Select
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725
' Category: advanced
' Scraped: 2025-10-09T09:05:59.777877

Sub Main
	Dim oADoc As AssemblyDocument = TryCast(ThisApplication.ActiveDocument, Inventor.AssemblyDocument)
	If oADoc Is Nothing Then
		Logger.Debug("Rule named '" & iLogicVb.RuleName & "' did not run because no assembly obtained!")
		Return
	End If
	
	Dim oSS As Inventor.SelectSet = oADoc.SelectSet
	Dim oOccs As ComponentOccurrences = oADoc.ComponentDefinition.Occurrences
	Dim oAllRefDocs As DocumentsEnumerator = oADoc.AllReferencedDocuments
		
	Dim oListOfPNs As List(Of String)
	oListOfPNs.Add("75430")
	oListOfPNs.Add("75424")

	For Each sPN As String In oListOfPNs
		'get the referenced Document with this Part Number (or FileName)
		For Each oRefDoc As Inventor.Document In oAllRefDocs
			Dim sPartNumber As String = oRefDoc.PropertySets.Item(3).Item(2).Value
			If sPartNumber = sPN Then
				Dim oRefOccs As ComponentOccurrencesEnumerator = oOccs.AllReferencedOccurrences(oRefDoc)
				If (oRefOccs Is Nothing) OrElse (oRefOccs.Count = 0) Then Continue For
				For Each oRefOcc As Inventor.ComponentOccurrence In oOccs
					'If oRefOcc.Name.StartsWith(sPN) Then
					oSS.Select(oRefOcc)
					'<<< DO SOMETHING WITH IT >>>
				Next 'oRefOcc
			End If
		Next 'oRefDoc
	Next 'sPN
	oADoc.Update2(True)
End Sub