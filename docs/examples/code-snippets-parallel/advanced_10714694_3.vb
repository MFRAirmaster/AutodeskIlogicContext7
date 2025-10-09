' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-09T08:50:38.589198

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
'define the variable "oName"
oName = "ModelState1"
For Each oOcc As ComponentOccurrence  In oADoc.ComponentDefinition.Occurrences
    If TypeOf oOcc.Definition Is PartComponentDefinition Then
		Dim oPDef As PartComponentDefinition = oOcc.Definition
		Dim oFound As Boolean  = False
		For Each oMS As ModelState In oPDef.ModelStates
			If oMS.Name = oName Then
				oFound = True
				Exit For 'found it, so don't create it
			End If
		Next
		If oFound = False Then
			'did not find it, so create it
			oPDef.ModelStates.Add(oName)
		End If
		oOcc.ActiveModelState = oName
	End If
Next