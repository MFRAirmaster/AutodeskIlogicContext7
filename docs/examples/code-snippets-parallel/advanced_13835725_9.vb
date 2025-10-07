' Title: Utilizing MakePath in SelectSet.Select
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:11:33.454183

Sub Main
	Dim oADoc As AssemblyDocument = TryCast(ThisDoc.Document, Inventor.AssemblyDocument)
	If oADoc Is Nothing Then
		Logger.Debug("The iLogic rule named '" & iLogicVb.RuleName & "' exited, because no AssemblyDocument was obtained.")
		Exit Sub
	End If
	
	'gather input from the user
	Dim sPartNumbersToFind As String = InputBox("Input part numbers separated by comma", "Select components")
	'if nothing entered, or is empty, exit the rule
	If String.IsNullOrWhiteSpace(sPartNumbersToFind) Then Return
	'attempt to put values into List
	If sPartNumbersToFind.Contains(",") Then
		oPartNumbersToFind = sPartNumbersToFind.Split(","c).ToList()
	Else
		If (oPartNumbersToFind Is Nothing) Then oPartNumbersToFind = New List(Of String)
		oPartNumbersToFind.Add(sPartNumbersToFind)
	End If
	'get assembly occurrrences collection
	Dim oOccs As ComponentOccurrences = oADoc.ComponentDefinition.Occurrences
	'call custom Sub routine to recurse this assemblies components
	'True means maintain context with parent assemblies
	RecurseComponents(oOccs, True, AddressOf ProcessComponent)
	
	'share the resulting List(Of ComponentOccurrence) object to Inventor's session memory
	SharedVariable.Value("FoundComponentsList") = oPartNumbersToFind
	'a later rule can retrieve this by: 
	'Dim FoundOccsList As List(Of ComponentOccurrence) = SharedVariable.Value("FoundComponentsList")
	
	'optionally update this assembly
	'oADoc.Update2(True)
	'optionally save this assembly
	'If oADoc.Dirty Then oADoc.Save2(True)
End Sub

'<<< declare 'global' variables here >>>
Dim oPartNumbersToFind As List(Of String)
Dim oFoundOccs As List(Of Inventor.ComponentOccurrence)

'just for assembly component recursion process (no need to customize)
Sub RecurseComponents(oComps As ComponentOccurrences, _
	preserveContext As Boolean, _
	ComponentProcess As Action(Of ComponentOccurrence))
	'validate input collection
	If (oComps Is Nothing) OrElse (oComps.Count = 0) Then Return
	'start iterating input collection
	For Each oComp As ComponentOccurrence In oComps
		'call custom task to run on this instance
		ComponentProcess(oComp)
		'if this occurrence is suppressed, we can not iterate through its sub occurrences
		If oComp.Suppressed Then Continue For
		'if this occurrrence is not an assembly, we do not want to iterate its sub occurrences
		If Not oComp.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then Continue For
		'attempt to recurse down into its sub-components, if possible, and it has some
		If preserveContext Then
			'<<< edits (can be) constrained to context of 'parent' assembly >>>
			RecurseComponents(oComp.SubOccurrences, preserveContext, ComponentProcess)
		Else
			'<<< edit each component's definition &/or referenced file independently >>>
			RecurseComponents(oComp.Definition.Occurrences, preserveContext, ComponentProcess)
		End If
	Next 'oComp
End Sub

'what to do to/with each occurrence
Sub ProcessComponent(oComp As ComponentOccurrence)
	If (oComp Is Nothing) OrElse oComp.Suppressed Then Return
	'skip virtual components
	'If (TypeOf oComp.Definition Is VirtualComponentDefinition) Then Return
	'skip welds within weldment type assemblies
	'If (TypeOf oComp.Definition Is WeldsComponentDefinition) Then Return
	'get referenced Document
	Dim oCompDoc As Inventor.Document = Nothing
	Try : oCompDoc = oComp.Definition.Document : Catch : End Try
	If oCompDoc Is Nothing Then Return
	'get Part Number iProperty value
	Dim sPN As String = oCompDoc.PropertySets.Item(3).Item(2).Value
	'if Part Number value is empty, then skip this component
	If String.IsNullOrWhiteSpace(sPN) Then Return
	'if this is 'NOT' one of the Part Number values we are looking for, then skip it
	If Not oPartNumbersToFind.Contains(sPN) Then Return
	'initialize the List, if necessary
	If (oFoundOccs Is Nothing) Then oFoundOccs = New List(Of Inventor.ComponentOccurrence)
	'meeds our requirements, so add it to our List
	If Not oFoundOccs.Contains(oComp) Then oFoundOccs.Add(oComp)
End Sub