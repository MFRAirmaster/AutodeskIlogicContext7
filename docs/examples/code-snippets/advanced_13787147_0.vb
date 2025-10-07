' Title: Assembly Place Component Preview
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/assembly-place-component-preview/td-p/13787147
' Category: advanced
' Scraped: 2025-10-07T13:15:16.402224

AddReference "MSInventorTools.dll"
Sub Main
If ThisApplication.UserInterfaceManager.ActiveEnvironment IsNot ThisApplication.UserInterfaceManager.Environments.Item("AMxAssemblyEnvironment") Then
	' Show an Error Message
	MessageBox.Show("This command can only be ran in the assembly environment", iLogicVb.RuleName, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
	Exit Sub
End If

' Create a MiniToolbar
Dim oMiniToolbar As MiniToolbar = ThisApplication.CommandManager.CreateMiniToolbar
	oMiniToolbar.ShowOK = True
	oMiniToolbar.ShowApply = True
	oMiniToolbar.ShowCancel = True
	oMiniToolbar.ShowOptionBox = False
	oMiniToolbar.ShowHandle = True

' Set a reference to the MiniToolbar Controls
Dim oMiniToolbarControls As MiniToolbarControls = oMiniToolbar.Controls

' Add a MiniToolbar Control to select a File
Dim oButton_SelectFile As MiniToolbarButton = oMiniToolbarControls.AddButton("Button_SelectFile", "Select File", "Select a file to place the component.", , )

' Add a MiniToolbar Control for the File Path
Dim oLabel_FilePath As MiniToolbarControl = oMiniToolbarControls.AddLabel("Label_FilePath", "", "", , )

' Add MiniToolbar Controls for the X, Y and Z Coordinates
Dim oNewLine_Coordinates As MiniToolbarControl = oMiniToolbarControls.AddNewLine

Dim oLabel_Coordinates As MiniToolbarControl = oMiniToolbarControls.AddLabel("Label_Coordinates", "Coordinates", "Enter the X, Y and Z coordinates.", , )

Dim oValueEditor_CoordinateX As MiniToolbarValueEditor = oMiniToolbarControls.AddValueEditor("ValueEditor_CoordinateX", "Enter the X coordinate", ValueUnitsTypeEnum.kLengthUnits, 0, "X", 0, , )
	oValueEditor_CoordinateX.AllowMeasure = False
	oValueEditor_CoordinateX.AllowShowDimensions = False
	
Dim oValueEditor_CoordinateY As MiniToolbarValueEditor = oMiniToolbarControls.AddValueEditor("ValueEditor_CoordinateY", "Enter the Y coordinate", ValueUnitsTypeEnum.kLengthUnits, 0, "Y", 0, , )
	oValueEditor_CoordinateY.AllowMeasure = False
	oValueEditor_CoordinateY.AllowShowDimensions = False
	
Dim oValueEditor_CoordinateZ As MiniToolbarValueEditor = oMiniToolbarControls.AddValueEditor("ValueEditor_CoordinateZ", "Enter the Z coordinate", ValueUnitsTypeEnum.kLengthUnits, 0, "Z", 0, , )
	oValueEditor_CoordinateZ.AllowMeasure = False
	oValueEditor_CoordinateZ.AllowShowDimensions = False

' Add MiniToolbar Controls for the X, Y and Z Rotations
Dim oNewLine_Rotations As MiniToolbarControl = oMiniToolbarControls.AddNewLine

Dim oLabel_Rotations As MiniToolbarControl = oMiniToolbarControls.AddLabel("Label_Rotations", "Rotations", "Enter the X, Y and Z Rotations.", , )

Dim ValueEditor_RotationX As MiniToolbarValueEditor = oMiniToolbarControls.AddValueEditor("ValueEditor_RotationX", "Enter the X Rotation", ValueUnitsTypeEnum.kAngleUnits, 0, "X", 0, , )
	ValueEditor_RotationX.AllowMeasure = False
	ValueEditor_RotationX.AllowShowDimensions = False
	
Dim oValueEditor_RotationY As MiniToolbarValueEditor = oMiniToolbarControls.AddValueEditor("ValueEditor_RotationY", "Enter the Y Rotation", ValueUnitsTypeEnum.kAngleUnits, 0, "Y", 0, , )
	oValueEditor_RotationY.AllowMeasure = False
	oValueEditor_RotationY.AllowShowDimensions = False
	
Dim oValueEditor_RotationZ As MiniToolbarValueEditor = oMiniToolbarControls.AddValueEditor("ValueEditor_RotationZ", "Enter the Z Rotation", ValueUnitsTypeEnum.kAngleUnits, 0, "Z", 0, , )
	oValueEditor_RotationZ.AllowMeasure = False
	oValueEditor_RotationZ.AllowShowDimensions = False

' Add MiniToolbar Controls for the Options
Dim oNewLine_Options As MiniToolbarControl = oMiniToolbarControls.AddNewLine

Dim oLabel_SelectInsertionPoint As MiniToolbarControl = oMiniToolbarControls.AddLabel("Label_SelectInsertionPoint", "Select Insertion Point", "Select the workpoint that should be placed on the specified coordinates.", , )

Dim oComboBox_SelectInsertionPoint As MiniToolbarComboBox = oMiniToolbarControls.AddComboBox("Combo_SelectInsertionPoint", True, True, 200)

' Disable the MiniToolbar Controls until a File is selected
EnableControls(oMiniToolbar, False)

' Show the MiniToolbar
oMiniToolbar.Visible = True

' Set a Reference to the MiniToolbar Events
Dim oMiniToolbarEvents As New MiniToolbarEvents(ThisApplication, oMiniToolbar)
End Sub

Public Shared Sub EnableControls(ByVal oMiniToolbar As MiniToolbar, ByVal bEnabled As Boolean)
	' Set a reference to the MiniToolbar Controls
	Dim oMiniToolbarControls As MiniToolbarControls = oMiniToolbar.Controls
	
	' Enable or Disable the MiniToolbar Controls
	oMiniToolbar.EnableOK = bEnabled
	oMiniToolbar.EnableApply = bEnabled
	
	Dim oLabel_FilePath As MiniToolbarControl = oMiniToolbarControls.Item("Label_FilePath")
		oLabel_FilePath.Visible = bEnabled
	
	Dim oLabel_Coordinates As MiniToolbarControl = oMiniToolbarControls.Item("Label_Coordinates")
		oLabel_Coordinates.Visible = bEnabled
	
	Dim oValueEditor_CoordinateX As MiniToolbarValueEditor = oMiniToolbarControls.Item("ValueEditor_CoordinateX")
		oValueEditor_CoordinateX.Visible = bEnabled
		
	Dim oValueEditor_CoordinateY As MiniToolbarValueEditor = oMiniToolbarControls.Item("ValueEditor_CoordinateY")
		oValueEditor_CoordinateY.Visible = bEnabled
		
	Dim oValueEditor_CoordinateZ As MiniToolbarValueEditor = oMiniToolbarControls.Item("ValueEditor_CoordinateZ")
		oValueEditor_CoordinateZ.Visible = bEnabled
	
	Dim oLabel_Rotations As MiniToolbarControl = oMiniToolbarControls.Item("Label_Rotations")
		oLabel_Rotations.Visible = bEnabled
	
	Dim ValueEditor_RotationX As MiniToolbarValueEditor = oMiniToolbarControls.Item("ValueEditor_RotationX")
		ValueEditor_RotationX.Visible = bEnabled
		
	Dim oValueEditor_RotationY As MiniToolbarValueEditor = oMiniToolbarControls.Item("ValueEditor_RotationY")
		oValueEditor_RotationY.Visible = bEnabled
		
	Dim oValueEditor_RotationZ As MiniToolbarValueEditor = oMiniToolbarControls.Item("ValueEditor_RotationZ")
		oValueEditor_RotationZ.Visible = bEnabled
	
	Dim oLabel_SelectInsertionPoint As MiniToolbarControl = oMiniToolbarControls.Item("Label_SelectInsertionPoint")
		oLabel_SelectInsertionPoint.Visible = bEnabled
	
	Dim oComboBox_SelectInsertionPoint As MiniToolbarComboBox = oMiniToolbarControls.Item("Combo_SelectInsertionPoint")
		oComboBox_SelectInsertionPoint.Visible = bEnabled
End Sub

Class MiniToolbarEvents
	Private ThisApplication As Inventor.Application
	Private WithEvents MiniToolbar As MiniToolbar
	Private WithEvents Button_SelectFile As MiniToolbarButton
	Private Label_FilePath As MiniToolbarControl
	Private ValueEditor_CoordinateX As MiniToolbarValueEditor
	Private ValueEditor_CoordinateY As MiniToolbarValueEditor
	Private ValueEditor_CoordinateZ As MiniToolbarValueEditor
	Private ValueEditorRotationX As MiniToolbarValueEditor
	Private ValueEditor_RotationY As MiniToolbarValueEditor
	Private ValueEditor_RotationZ As MiniToolbarValueEditor
	Private Label_SelectInsertionPoint As MiniToolbarControl
	Private ComboBox_SelectInsertionPoint As MiniToolbarComboBox
	Private FilePath As String = ""
	
	Public Sub New(ByVal oApplication As Inventor.Application, ByVal oMiniToolbar As MiniToolbar)
		ThisApplication = oApplication
		MiniToolbar = oMiniToolbar
	    Button_SelectFile = MiniToolbar.Controls.Item("Button_SelectFile")
		Label_FilePath = MiniToolbar.Controls.Item("Label_FilePath")
		ValueEditor_CoordinateX = MiniToolbar.Controls.Item("ValueEditor_CoordinateX")
		ValueEditor_CoordinateY = MiniToolbar.Controls.Item("ValueEditor_CoordinateY")
		ValueEditor_CoordinateZ = MiniToolbar.Controls.Item("ValueEditor_CoordinateZ")
		ValueEditorRotationX = MiniToolbar.Controls.Item("ValueEditor_RotationX")
		ValueEditor_RotationY = MiniToolbar.Controls.Item("ValueEditor_RotationY")
		ValueEditor_RotationZ = MiniToolbar.Controls.Item("ValueEditor_RotationZ")
		Label_SelectInsertionPoint = MiniToolbar.Controls.Item("Label_SelectInsertionPoint")
		ComboBox_SelectInsertionPoint = MiniToolbar.Controls.Item("Combo_SelectInsertionPoint")
	End Sub
	
	Private Sub Button_SelectFile_OnClick() Handles Button_SelectFile.OnClick
		' Select a Model File
		FilePath = MSInventorTools.MS_OpenFileDialogs.OpenModelFilesDialog(False)(0)
		
		' Enable or Disable the MiniToolbar Controls
		If String.IsNullOrWhiteSpace(FilePath) = False Then
			Label_FilePath.DisplayName = FilePath
			ThisRule.EnableControls(MiniToolbar, True)
			GetWorkPointsFromComponent()
		Else
			Label_FilePath.DisplayName = ""
			ComboBox_SelectInsertionPoint.Clear
			ThisRule.EnableControls(MiniToolbar, False)
		End If
	End Sub
	
	Private Sub MiniToolbar_OnOK() Handles MiniToolbar.OnOK 
		PlaceComponent()
	End Sub
	
	Private Sub MiniToolbar_OnApply() Handles MiniToolbar.OnApply
		PlaceComponent()
	End Sub
	
	Private Sub MiniToolbar_OnCancel() Handles MiniToolbar.OnCancel
	End Sub
	
	Private Sub GetWorkPointsFromComponent()
		' Open the Document
		Dim oDoc As Document = ThisApplication.Documents.Open(FilePath, False)
		
		' Set a reference to the Component Definition
		Dim oCompDef As ComponentDefinition = oDoc.ComponentDefinition
		
		' Set a reference to the Work Points Collection
		Dim oWorkPoints As WorkPoints = oCompDef.WorkPoints
		
		' Clear the Combo Box
		ComboBox_SelectInsertionPoint.Clear
		
		' Populate the Combo Box
		For i As Integer = 1 To oWorkPoints.Count
			ComboBox_SelectInsertionPoint.AddItem(oWorkPoints.Item(i).Name, "")
		Next
		
		' Close the Document
		oDoc.Close(True)
	End Sub
	
	Private Sub PlaceComponent()
		' Check if the expressions are valid
		If ValueEditor_CoordinateX.IsExpressionValid = True AndAlso
				ValueEditor_CoordinateY.IsExpressionValid = True AndAlso
				ValueEditor_CoordinateZ.IsExpressionValid = True AndAlso
				ValueEditorRotationX.IsExpressionValid = True AndAlso
				ValueEditor_RotationY.IsExpressionValid = True AndAlso
				ValueEditor_RotationZ.IsExpressionValid = True Then
			
			' Set the Values
			Dim oUnitsOfMeasure As UnitsOfMeasure = ThisApplication.UnitsOfMeasure
			Dim LocationX As Double = oUnitsOfMeasure.ConvertUnits(ValueEditor_CoordinateX.Value, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)
			Dim LocationY As Double = oUnitsOfMeasure.ConvertUnits(ValueEditor_CoordinateY.Value, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)
			Dim LocationZ As Double = oUnitsOfMeasure.ConvertUnits(ValueEditor_CoordinateZ.Value, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)
			Dim RotationX As Double = oUnitsOfMeasure.ConvertUnits(ValueEditorRotationX.Value, UnitsTypeEnum.kDatabaseAngleUnits, UnitsTypeEnum.kDegreeAngleUnits)
			Dim RotationY As Double = oUnitsOfMeasure.ConvertUnits(ValueEditor_RotationY.Value, UnitsTypeEnum.kDatabaseAngleUnits, UnitsTypeEnum.kDegreeAngleUnits)
			Dim RotationZ As Double = oUnitsOfMeasure.ConvertUnits(ValueEditor_RotationZ.Value, UnitsTypeEnum.kDatabaseAngleUnits, UnitsTypeEnum.kDegreeAngleUnits)
			
			' Place the Component
			Dim oOcc As ComponentOccurrence = MSInventorTools.MS_Occurrence.PlaceOccurrence(FilePath, LocationX, LocationY, LocationZ, RotationX, RotationY, RotationZ, 0, 0, 0, 0, 0, 0, True)
			
			' Move the Occurrence to its Insertion Point
			MSInventorTools.MS_Occurrence.MoveOccurrenceToInsertionPoint(oOcc, ComboBox_SelectInsertionPoint.SelectedItem.Text)
		End If
	End Sub
End Class