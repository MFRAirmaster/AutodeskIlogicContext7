' Title: CircularPatternOccurrence
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/circularpatternoccurrence/td-p/13835553
' Category: advanced
' Scraped: 2025-10-09T09:02:38.376983

'get the current document, then try to Cast its Type to the AssemblyDocument Type
Dim oADoc As AssemblyDocument = TryCast(ThisDoc.Document, Inventor.AssemblyDocument)
'if that did not work, then we did not get a value, so exit the rule
If oADoc Is Nothing Then Return
'get the assembly's component definition (where all geometry and coordinate system are)
Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
'get just the occurrence patterns collection
'thic can contain rectangular, circular, and feature-based patterns
Dim oOccPatts As OccurrencePatterns = oADef.OccurrencePatterns
'if there are none, then exit the eule
If oOccPatts.Count = 0 Then Return
'start iterating through each OccurrencePattern
For Each oOccPatt As OccurrencePattern In oOccPatts
	'if it is suppressed, then skip over it (because we can't do anything with it)
	If oOccPatt.Suppressed Then Continue For
	'if it is not a 'circular' type, then skip over it
	If (Not TypeOf oOccPatt Is CircularOccurrencePattern) Then Continue For
	'declare a variable to hold the CircularOccurrencePattern Type object
	Dim oCircOccPatt As CircularOccurrencePattern = oOccPatt
	'check if its name matches what we want to get
	'<<< CHANGE THIS AS NEEDED >>>
	If oCircOccPatt.Name = "My Circular Occurrence Pattern" Then
		Dim oOccPattElmts As OccurrencePatternElements = oCircOccPatt.OccurrencePatternElements
		'get the 'Last' element in the pattern
		Dim oOccPattElmt As OccurrencePatternElement = oOccPattElmts.Item(oOccPattElmts.Count)
		'get first component occurrence in this element
		Dim oPattElmtOcc As ComponentOccurrence = oOccPattElmt.Occurrences.Item(1)
		'<<< CHANGE THIS AS NEEDED >>>
		Dim sDVR_Name As String = "Specific Design View Representation Name"
		Try
			'try so set the DVR of this component occurrence to a specific one available in its referenced document
			oPattElmtOcc.SetDesignViewRepresentation(sDVR_Name, , True)
		Catch
			Logger.Error("Failed to set View Representation of circular pattern element component occurrence!")
		End Try
	End If
Next 'oOccPatt
'update the main assembly document
oADoc.Update2(True)