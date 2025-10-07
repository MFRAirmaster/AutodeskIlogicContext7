' Title: what is the &quot;correct&quot; way to write this line of code for a rule?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/what-is-the-quot-correct-quot-way-to-write-this-line-of-code-for/td-p/13761867#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:23:27.340982

Dim oAssyDoc As AssemblyDocument
oAssyDoc = ThisApplication.ActiveDocument

Dim oAssyCompDef As AssemblyComponentDefinition
oAssyCompDef = oAssyDoc.ComponentDefinition

Dim oOccurrences As ComponentOccurrences
oOccurrences = oAssyCompDef.Occurrences

Dim oLeafOccs As ComponentOccurrencesEnumerator
oLeafOccs = oOccurrences.AllLeafOccurrences

Dim oLeafOcc As ComponentOccurrence

For Each oLeafOcc In oLeafOccs
 ' DO YOUR WORK HERE
	
' ########## - TURN OFF VISIBILITY FOR ALL PARTS CONTAINING STRING BELOW - ###########	
	Dim oPDef As PartComponentDefinition = oLeafOcc.Definition
	oDVRepName = "Hidden"
	oDVReps = oPDef.RepresentationsManager.DesignViewRepresentations
	For Each oDVRep As DesignViewRepresentation In oDVReps
		If oDVRep.Name = oDVRepName Then
			'Found it'
			oLeafOcc.SetDesignViewRepresentation(oDVRepName, , True)
		End If
		Next
' ########## - END OF MY SECTION - ###########	

Next