' Title: Utilizing MakePath in SelectSet.Select
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:33:22.602874

Sub Main
	oDoc = ThisDoc.Document
	
	Dim oComponentsString As String = InputBox("Input part numbers separated by comma", "Select components")
	
	oComponentsList = oComponentsString.Split(","c).ToList
	
	oObjCol = ThisApplication.TransientObjects.CreateObjectCollection
	assemblyRunner(oDoc)
	
	oDoc.SelectSet.SelectMultiple(oObjCol)
End Sub

Dim oDoc As AssemblyDocument
Dim oComponentsList As List(Of String)
Dim alreadyInList As Boolean
Dim occList As New List(Of ComponentOccurrence)
Dim oObjCol As ObjectCollection

Sub assemblyRunner(aDoc As AssemblyDocument)
	
	For Each oOcc In aDoc.ComponentDefinition.Occurrences
		Dim oOccDoc As Document = oOcc.Definition.Document
		Dim oPartNo As String = oOccDoc.PropertySets.Item(3).Item("Part Number").Value
		
		If oComponentsList.Contains(oPartNo) Then
			For Each comp In occList
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
				occList.Add(oOcc)
				oObjCol.Add(oOcc)
			End If
		End If
		If oOccDoc.DocumentType = kAssemblyDocumentObject Then
			assemblyRunner(oOccDoc)
		End If
	Next
End Sub