' Title: Utilizing MakePath in SelectSet.Select
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725
' Category: advanced
' Scraped: 2025-10-09T09:05:59.777877

Sub Main
	oDoc = ThisDoc.Document
	
	Dim oComponentsString As String = InputBox("Insert part numbers, separated by comma", "Select components")
	
	oComponentsList = oComponentsString.Split(","c).ToList
	
	oObjCol = ThisApplication.TransientObjects.CreateObjectCollection
	
	' On the first level, run through all occurrences, and add those found in the list to the object collection
	firstLevelRunner(oDoc)
	
	oDoc.SelectSet.SelectMultiple(oObjCol)
End Sub

Dim oDoc As AssemblyDocument
Dim oComponentsList As List(Of String)
Dim alreadyOnList As Boolean
Dim aOcc As ComponentOccurrence
Dim oOccProxy As ComponentOccurrenceProxy
Dim oPartNo As String
Dim oObjCol As ObjectCollection

Sub firstLevelRunner(aDoc As AssemblyDocument)
	Dim aOccDoc As Document
	
	' Loop through all first level occurrences, ignoring those already found using the "alreadyOnList"-variable
	For Each aOcc As ComponentOccurrence In aDoc.ComponentDefinition.Occurrences
		aOccDoc = aOcc.Definition.Document
		oPartNo = aOccDoc.PropertySets.Item(3).Item("Part Number").Value
		
		If oComponentsList.Contains(oPartNo) Then
			For Each comp In oObjCol
				compPartNo = comp.Definition.Document.PropertySets.Item(3).Item("Part Number").Value
				If compPartNo = oPartNo Then
					alreadyOnList = True
					Exit For
				End If
			Next
			If alreadyOnList Then
				alreadyOnList = False
				Continue For
			Else
				oObjCol.Add(aOcc)
			End If
		End If
		
		' If the current occurence is an assembly, loop through its lower levels using occurrence proxies
		If aOccDoc.DocumentType = kAssemblyDocumentObject Then
			occRunner(aOcc)
		End If
	Next
End Sub

Sub occRunner(oOcc As ComponentOccurrence)
	Dim occProxyDoc As Document
	
	For Each occProxy As ComponentOccurrenceProxy In oOcc.SubOccurrences
		occProxyDoc = occProxy.Definition.Document
		oPartNo = occProxyDoc.PropertySets.Item(3).Item("Part Number").Value
		
		If oComponentsList.Contains(oPartNo) Then
			For Each comp In oObjCol
				compPartNo = comp.Definition.Document.PropertySets.Item(3).Item("Part Number").Value
				If compPartNo = oPartNo Then
					alreadyOnList = True
					Exit For
				End If
			Next
			If alreadyOnList Then
				alreadyOnList = False
				Continue For
			Else
				oObjCol.Add(occProxy)
			End If
		End If
		If occProxyDoc.DocumentType = kAssemblyDocumentObject Then
			occRunner(occProxy)
		End If
	Next
End Sub