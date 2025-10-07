' Title: ilogic rule to create a rectangular pattern in an assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-create-a-rectangular-pattern-in-an-assembly/td-p/9325882
' Category: advanced
' Scraped: 2025-10-07T14:03:29.616764

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MessageBox.Show("THIS RULE '" & iLogicVb.RuleName & "' ONLY WORKS FOR ASSEMBLY DOCUMENTS.", "WRONG DOCUMENT TYPE", MessageBoxButtons.OK, MessageBoxIcon.Stop)
	Return
End If

Dim oADoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
Dim oOccs As ComponentOccurrences = oADef.Occurrences
Dim oUParams As UserParameters = oADef.Parameters.UserParameters

Try
	Dim oPatt As RectangularOccurrencePattern = oADef.OccurrencePatterns.Item("PADR_PARAFUSOS_INOX")
	Dim oDelete As MsgBoxResult = MsgBox("An OccurrencePattern named '" & oPatt.Name & "' already exists." & vbNewLine &
	"Do you want to Delete it and its Components?", vbYesNo + vbQuestion, "PATTERN EXISTS - DELETE?")
	If oDelete = vbYes Then
		'iLogicVb.RunExternalRule("Name Of Rule To Delete Them")
		'Or
		oPatt.Delete
		'If the following names don't match actual items exactly, it will throw an error.
'		Dim oOcc1 As ComponentOccurrence = oOccs.ItemByName("4349:1")
'		Dim oOcc2 As ComponentOccurrence = oOccs.ItemByName("0157:1")
'		oOcc1.Delete
'		oOcc2.Delete
		'This can become difficult when more than one of the same item is in the same assembly.
		'Instead you can use a loop code which searches for a name that 'contains' some text, like this
		For Each oOcc As ComponentOccurrence In oOccs
			If oOcc.Name.Contains("4349") Or oOcc.Name.Contains("0157") Then
				oOcc.Delete
			End If
		Next				
	End If
	
	Dim oEdit As MsgBoxResult = MsgBox("Do you want to Edit its Properties or Values?", vbYesNo + vbQuestion, "EDIT IT?")
	If oEdit = vbYes Then
		'Perhaps use a series of InputBox or InputListBox lines of code here to offer the user the ability to change its values
		'Something like the following
		'ColumnCount is Read Only, so to change its value, change the value of the Parameter being used by it
		Dim oColumnCount As Parameter = oPatt.ColumnCount
		oColumnCount.Value = InputBox("Enter New Column Count.")
		Dim oAxisList As List(Of String)
		oAxisList.Add("X Axis")
		oAxisList.Add("Y Axis")
		oAxisList.Add("Z Axis")
		oPatt.ColumnEntity = InputListBox("Choose Axis For Column Entity.", oAxisList, oAxisList.Item(1), "COLUMN ENTITY")
		oPatt.ColumnEntityNaturalDirection = InputRadioBox("Use Column Entity Natural Direction?", True, False)
		'ColumnOffset is Read Only, so to change its value, change the value of the Parameter being used by it
		Dim oColumnOffset As Parameter = oPatt.ColumnOffset
		oColumnOffset.Value = InputBox("Enter New Column Offset.")
		'Etc
	ElseIf oEdit = vbNo Then
		Return
	End If
Catch
	'It doesn't exist yet, so create it
	Dim oTO As TransientObjects = ThisApplication.TransientObjects
	Dim oObjects As ObjectCollection = oTO.CreateObjectCollection
	oObjects.Add(oOccs.ItemByName("4349:1"))
	oObjects.Add(oOccs.ItemByName("0157:1"))
	Dim oRowEntity As Object = oADef.WorkAxes.Item("Z Axis")
	Dim oColumnEntity As Object = oADef.WorkAxes.Item("Y Axis")
	'You may also want to include a few lines here which check to make sure these UserParameters exist.
'	Dim oColumnOffset As UserParameter = 
'	Dim oColumnCount As UserParameter = 
	Dim oRowOffset As UserParameter = oUParams.Item("DIST_SUP")
	Dim oRowCount As UserParameter = oUParams.Item("QTD_SUP")

	Dim oPatt As RectangularOccurrencePattern
	oPatt = oADef.OccurrencePatterns.AddRectangularPattern(oObjects, oColumnEntity, True, 0, 1, oRowEntity, True, oRowOffset, oRowCount)
	oPatt.Name = "PADR_PARAFUSOS_INOX"
End Try