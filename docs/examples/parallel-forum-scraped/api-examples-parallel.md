# Parallel Api iLogic Examples

Generated from 80 forum posts with advanced parallel scraping.

**Generated:** 2025-10-09T09:10:07.538400

---

## how to change weld size in weld symbol by ilogic

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/how-to-change-weld-size-in-weld-symbol-by-ilogic/td-p/13788475](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-change-weld-size-in-weld-symbol-by-ilogic/td-p/13788475)

**Author:** sameer_maratheRNRKJ

**Date:** ‎08-29-2025
	
		
		04:13 AM

**Code:**

```vb
Dim obj = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Select a weld symbol")

If (TypeOf obj IsNot DrawingWeldingSymbol) Then
    MsgBox("You did not select a welding symbol")
    Return
End If

Dim symbol As DrawingWeldingSymbol = obj
Dim def As DrawingWeldingSymbolDefinition = symbol.Definitions.Item(1)
Dim symbolDef1 = def.WeldSymbolOne
Dim symbolDef2 = def.WeldSymbolTwo

symbolDef1.Prefix = "Prefix 1"
symbolDef1.Leg1 = "Leg 1.1"
symbolDef1.Leg2 = "Leg 1.2"
symbolDef1.WeldSymbolType = WeldSymbolTypeEnum.kFilletWeldSymbolType
symbolDef1.Length = "Length 1"
symbolDef1.Pitch = "Pitch 1"

symbolDef2.Prefix = "Prefix 2"
symbolDef2.Leg1 = "Leg 2.1"
symbolDef2.Leg2 = "Leg 2.2"
symbolDef2.WeldSymbolType = WeldSymbolTypeEnum.kFilletWeldSymbolType
symbolDef2.Length = "Length 2"
symbolDef2.Pitch = "Pitch 2"
```

```vb
Dim obj = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Select a weld symbol")

If (TypeOf obj IsNot DrawingWeldingSymbol) Then
    MsgBox("You did not select a welding symbol")
    Return
End If

Dim symbol As DrawingWeldingSymbol = obj
Dim def As DrawingWeldingSymbolDefinition = symbol.Definitions.Item(1)
Dim symbolDef1 = def.WeldSymbolOne
Dim symbolDef2 = def.WeldSymbolTwo

symbolDef1.Prefix = "Prefix 1"
symbolDef1.Leg1 = "Leg 1.1"
symbolDef1.Leg2 = "Leg 1.2"
symbolDef1.WeldSymbolType = WeldSymbolTypeEnum.kFilletWeldSymbolType
symbolDef1.Length = "Length 1"
symbolDef1.Pitch = "Pitch 1"

symbolDef2.Prefix = "Prefix 2"
symbolDef2.Leg1 = "Leg 2.1"
symbolDef2.Leg2 = "Leg 2.2"
symbolDef2.WeldSymbolType = WeldSymbolTypeEnum.kFilletWeldSymbolType
symbolDef2.Length = "Length 2"
symbolDef2.Pitch = "Pitch 2"
```

---

## Inventor 2026 Inventor Application, GetActiveObject is not a member of Marshal

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2026-inventor-application-getactiveobject-is-not-a/td-p/13830311](https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2026-inventor-application-getactiveobject-is-not-a/td-p/13830311)

**Author:** A.Acheson

**Date:** ‎09-29-2025
	
		
		02:31 PM

**Description:** I am having trouble converting a VB.Net form designed to run in the ilogic environment. Error on Line 56 : 'GetActiveObject' is not a member of 'Marshal'. I am not using visual studio so have no access to how the application object is created.  Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	 ...

**Code:**

```vb
Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	    'Note: if you shut down the Inventor session that was started
	    'this(way) there is still an Inventor.exe running. We will use
	    'this Boolean to test whether or not the Inventor App  will
	    'need to be shut down.
	   Catch ex2 As Exception
	     MsgBox(ex2.ToString())
	     MsgBox("Unable to get or start Inventor")
	   End Try
End Try
```

```vb
Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	    'Note: if you shut down the Inventor session that was started
	    'this(way) there is still an Inventor.exe running. We will use
	    'this Boolean to test whether or not the Inventor App  will
	    'need to be shut down.
	   Catch ex2 As Exception
	     MsgBox(ex2.ToString())
	     MsgBox("Unable to get or start Inventor")
	   End Try
End Try
```

```vb
_invApp = Thisapplication
```

```vb
_invApp = Thisapplication
```

```vb
' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
			'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
			'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
            'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub
```

```vb
' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
			'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
			'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
            'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub
```

```vb
_invApp.Visible = True
```

```vb
_invApp.Visible = True
```

```vb
ThisApplication.Visible = True
```

```vb
ThisApplication.Visible = True
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
```

```vb
Public Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	Public thisApp As Inventor.Application
	Public Sub New
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub

		
End Class

Sub Main()
 Try
	thisApp = ThisApplication
	oDrawDoc = ThisApplication.ActiveDocument
	MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
  Catch ex As Exception
  End Try

	Dim oMyForm As New WinForm 
	oMyForm.Show

End Sub
```

```vb
Public Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	Public thisApp As Inventor.Application
	Public Sub New
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub

		
End Class

Sub Main()
 Try
	thisApp = ThisApplication
	oDrawDoc = ThisApplication.ActiveDocument
	MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
  Catch ex As Exception
  End Try

	Dim oMyForm As New WinForm 
	oMyForm.Show

End Sub
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
 

 

 Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	'Public  thisApp As Inventor.Application=ThisApplication
	Public Sub New(thisApp As Inventor.Application)
		oDrawDoc = thisApp.ActiveDocument
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		Me.Close
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub


		
End Class

Sub Main()


	Dim oMyForm As New WinForm (ThisApplication)
	oMyForm.Show

End Sub
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
 

 

 Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	'Public  thisApp As Inventor.Application=ThisApplication
	Public Sub New(thisApp As Inventor.Application)
		oDrawDoc = thisApp.ActiveDocument
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		Me.Close
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub


		
End Class

Sub Main()


	Dim oMyForm As New WinForm (ThisApplication)
	oMyForm.Show

End Sub
```

---

## problem with rules on remote pc

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/problem-with-rules-on-remote-pc/td-p/13810727](https://forums.autodesk.com/t5/inventor-programming-forum/problem-with-rules-on-remote-pc/td-p/13810727)

**Author:** Frank_DoucetLYPFL

**Date:** ‎09-15-2025
	
		
		02:08 AM

**Description:** i made a rules in inventor for copying and renaming all files from one folder to another, at home it worked like it schould, at work we use remote pc and there i did not work, getting errorLine29 : 'directory' is not declared, it may be inaccessible due to its protection levelLine30 : 'directory' is not declared, it may be inaccessible due to its protection levelLine34 : 'directory' is not declared, it may be inaccessible due to its protection levelLine40 : 'path' is not declared, it may be inac...

**Code:**

```vb
Imports System.IO

Sub Main()
' Define source folder
Dim sourceFolder As String = "C:\Users\Frank\Documents\Inventor\ilogic Afsluiter"

' Ask user to choose destination folder
Dim fbd As New FolderBrowserDialog()
fbd.Description = "Select the destination folder"
fbd.ShowNewFolderButton = True

Dim result As DialogResult = fbd.ShowDialog()
If result <> DialogResult.OK Then
MessageBox.Show("Operation cancelled.", "iLogic")
Exit Sub
End If

Dim destFolder As String = fbd.SelectedPath

' Ask user for prefix (e.g. M593)
Dim prefix As String = InputBox("Enter the prefix (example: M593)", "File Prefix", "M593")
If prefix = "" Then
MessageBox.Show("No prefix entered. Operation cancelled.", "iLogic")
Exit Sub
End If

' Create destination folder if it doesn’t exist
If Not Directory.Exists(destFolder) Then
Directory.CreateDirectory(destFolder)
End If

' Get all files in source folder
Dim files() As String = Directory.GetFiles(sourceFolder)

' Counter for renaming
Dim counter As Integer = 0

For Each filePath As String In files
Dim ext As String = Path.GetExtension(filePath)
Dim newName As String = prefix & "-" & counter.ToString("000") & "-00" & ext
Dim destPath As String = Path.Combine(destFolder, newName)

' Copy file (overwrite if already exists)
File.Copy(filePath, destPath, True)

counter += 1
Next

MessageBox.Show("Files copied and renamed successfully to:" & vbCrLf & destFolder, "iLogic")
End Sub
```

```vb
Imports System.IO

Sub Main()
' Define source folder
Dim sourceFolder As String = "C:\Users\Frank\Documents\Inventor\ilogic Afsluiter"

' Ask user to choose destination folder
Dim fbd As New FolderBrowserDialog()
fbd.Description = "Select the destination folder"
fbd.ShowNewFolderButton = True

Dim result As DialogResult = fbd.ShowDialog()
If result <> DialogResult.OK Then
MessageBox.Show("Operation cancelled.", "iLogic")
Exit Sub
End If

Dim destFolder As String = fbd.SelectedPath

' Ask user for prefix (e.g. M593)
Dim prefix As String = InputBox("Enter the prefix (example: M593)", "File Prefix", "M593")
If prefix = "" Then
MessageBox.Show("No prefix entered. Operation cancelled.", "iLogic")
Exit Sub
End If

' Create destination folder if it doesn’t exist
If Not Directory.Exists(destFolder) Then
Directory.CreateDirectory(destFolder)
End If

' Get all files in source folder
Dim files() As String = Directory.GetFiles(sourceFolder)

' Counter for renaming
Dim counter As Integer = 0

For Each filePath As String In files
Dim ext As String = Path.GetExtension(filePath)
Dim newName As String = prefix & "-" & counter.ToString("000") & "-00" & ext
Dim destPath As String = Path.Combine(destFolder, newName)

' Copy file (overwrite if already exists)
File.Copy(filePath, destPath, True)

counter += 1
Next

MessageBox.Show("Files copied and renamed successfully to:" & vbCrLf & destFolder, "iLogic")
End Sub
```

---

## iLogic - Check active project is true

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-check-active-project-is-true/td-p/10218046#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-check-active-project-is-true/td-p/10218046#messageview_0)

**Author:** neil.cooke

**Date:** ‎04-07-2021
	
		
		02:19 AM

**Description:** I have found similar requests with solution but I am quite a new user to iLogic and perhaps missing the obvious (Sorry). I am trying to check that a specific project file is active, if it is not true then to pop up a message box stating you need to set the active project to the specified file. Please could someone help. Many thanks 
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Public Sub InventorProjectFile(oInvapp As Inventor.Application)

        Dim proj As DesignProject

        Try
            proj = oInvapp.DesignProjectManager.DesignProjects.ItemByName("C:\#######.ipj")
            proj.Activate()
        Catch ex As Exception
            If Not oInvapp.DesignProjectManager.ActiveDesignProject.FullFileName = "C:\#######.ipj" Then
                MsgBox("The wrong Project file has been Loaded !" & vbCrLf & "The Current Project file = " & oInvapp.DesignProjectManager.ActiveDesignProject.Name & vbCrLf & "Load the " & "C:\#######.ipj" & " file", vbCritical)
            End If
        End Try
end sub
```

```vb
Public Sub InventorProjectFile(oInvapp As Inventor.Application)

        Dim proj As DesignProject

        Try
            proj = oInvapp.DesignProjectManager.DesignProjects.ItemByName("C:\#######.ipj")
            proj.Activate()
        Catch ex As Exception
            If Not oInvapp.DesignProjectManager.ActiveDesignProject.FullFileName = "C:\#######.ipj" Then
                MsgBox("The wrong Project file has been Loaded !" & vbCrLf & "The Current Project file = " & oInvapp.DesignProjectManager.ActiveDesignProject.Name & vbCrLf & "Load the " & "C:\#######.ipj" & " file", vbCritical)
            End If
        End Try
end sub
```

```vb
Sub Main
	Const sProjPath As String = "C:\Temp\"
	Const sProjFileName As String = "MyInventorProject.ipj"
	Const sProjFile As String = "C:\Temp\MyInventorProject.ipj"
	Dim oProjMgr As Inventor.DesignProjectManager = ThisApplication.DesignProjectManager
	'get FullFileName of active DesignProject
	Dim sActiveProj As String = oProjMgr.ActiveDesignProject.FullFileName
	'only react if it is not the one expected / wanted
	If Not sActiveProj = sProjFile Then
		'Log it, or let user know
		Logger.Info("Active Project File Is:  " & sActiveProj)
		'MessageBox.Show("Active Project File Is:  " & sActiveProj)
		'try to find / get the DesignProject we want, by its FullFileName
		Dim oMyProj As Inventor.DesignProject = Nothing
		Try
			oMyProj = oProjMgr.DesignProjects.ItemByName(sProjFile)
		Catch
			Logger.Warn("Specified Inventor DesignProject Not Fount In Projects Collection!")
		End Try
		'if it was not found in that collection, then add it
		If oMyProj Is Nothing Then
			Try
				oMyProj = oProjMgr.DesignProjects.AddExisting(sProjFile)
			Catch
				Logger.Error("Error Adding Specified DesignProject To Projects Collection!")
			End Try
		End If
		'if we now have it, then activate it
		If oMyProj IsNot Nothing Then
			Try
				'True = Set as Default Project
				oMyProj.Activate(True)
			Catch
				Logger.Error("Error Activating Specified DesignProject!")
			End Try
		End If
	End If
End Sub
```

```vb
Sub Main
	Const sProjPath As String = "C:\Temp\"
	Const sProjFileName As String = "MyInventorProject.ipj"
	Const sProjFile As String = "C:\Temp\MyInventorProject.ipj"
	Dim oProjMgr As Inventor.DesignProjectManager = ThisApplication.DesignProjectManager
	'get FullFileName of active DesignProject
	Dim sActiveProj As String = oProjMgr.ActiveDesignProject.FullFileName
	'only react if it is not the one expected / wanted
	If Not sActiveProj = sProjFile Then
		'Log it, or let user know
		Logger.Info("Active Project File Is:  " & sActiveProj)
		'MessageBox.Show("Active Project File Is:  " & sActiveProj)
		'try to find / get the DesignProject we want, by its FullFileName
		Dim oMyProj As Inventor.DesignProject = Nothing
		Try
			oMyProj = oProjMgr.DesignProjects.ItemByName(sProjFile)
		Catch
			Logger.Warn("Specified Inventor DesignProject Not Fount In Projects Collection!")
		End Try
		'if it was not found in that collection, then add it
		If oMyProj Is Nothing Then
			Try
				oMyProj = oProjMgr.DesignProjects.AddExisting(sProjFile)
			Catch
				Logger.Error("Error Adding Specified DesignProject To Projects Collection!")
			End Try
		End If
		'if we now have it, then activate it
		If oMyProj IsNot Nothing Then
			Try
				'True = Set as Default Project
				oMyProj.Activate(True)
			Catch
				Logger.Error("Error Activating Specified DesignProject!")
			End Try
		End If
	End If
End Sub
```

---

## Looking for some &quot;fun&quot; iLogic code for a T-Shirt

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/looking-for-some-quot-fun-quot-ilogic-code-for-a-t-shirt/td-p/13840538](https://forums.autodesk.com/t5/inventor-programming-forum/looking-for-some-quot-fun-quot-ilogic-code-for-a-t-shirt/td-p/13840538)

**Author:** chris

**Date:** ‎10-06-2025
	
		
		07:08 PM

**Description:** I'm looking to add some fun or unique iLogic code for a T-Shirt that I'm making for the designers at work. Can anyone think of some fun code that might look cool or fun on the back of a shirt? (it can be short or long), just looking for something fun/cool, I plan on using the code edit colors as wellExample: iLogic code to create a cube or sphere, I though about some iProperties, but I'm open to ideas!

**Code:**

```vb
Dim Stress As Double = 0
While Project.Deadline < Now
    Stress += 1
    Sleep(1)
End While
```

```vb
Do Until Sketch.Perfect = True
    IterateDesign()
Loop
```

```vb
Try
    OpenInventor()
    StartDesigning("GeniusPart")
    While Not Done
        Iterate()
        Overthink()
    End While
Catch ex As BurnoutException
    Coffee.Refill()
    Resume Next
End Try
```

```vb
If Coffee.Level < 10 Then
    Call Coffee.Refill()
    Return
End If
Call Design.Part("Brilliance")
```

```vb
If You.AreConstrainedTo(Me) Then
    Life.IsFullyDefined = True
End If
```

```vb
If You.Fit(My.Mate) Then
    Let'sAssemble()
End If
```

```vb
If Not Coffee.Exists Then
    System.Collapse()
End If
```

```vb
Sub EngineeringProcess()
    ' Engineering is spelled: C-H-A-N-G-E
    
    Design = CreateInitialConcept()
    SendToClient()
    
    While (ProjectNotCancelled)
        changeRequest = Client.SendEmail()
        
        Select Case changeRequest.Severity
            Case "Minor tweak"
                RedoEverything()
            Case "Small adjustment"
                StartFromScratch()
            Case "Just one little thing"
                QuestionCareerPath()
        End Select
        
        UpdateRevisionNumber() ' Now at Rev ZZ.47
        
        If (design.LooksLikeOriginalConcept = False) Then
            Console.WriteLine("Nailed it! ��")
        End If
        
        WaitForNextChangeRequest(milliseconds:=3)
    End While
    
    ' Note: ProjectNotCancelled always returns True
End Sub
```

---

## Decimal places in iLogic form

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/decimal-places-in-ilogic-form/td-p/13806789](https://forums.autodesk.com/t5/inventor-programming-forum/decimal-places-in-ilogic-form/td-p/13806789)

**Author:** Curtis_Waguespack

**Date:** ‎09-11-2025
	
		
		10:33 AM

**Description:** Problem with decimal places in iLogic form in Inventor 2025. See attached files.
 
Here is the issue:


 
 
I thought I knew how to control this, by handing the value to a variable in a rule, rounding it, and then pushing that rounded value back out to the parameter, like this.

But I'm not getting that to work, now. Not sure if something changed, or...Does anyone know how to fix this?
 
Note that some increments display as expected

 
others do not

 
 
 
					
				
			
			
				
	

			
			
				...

**Code:**

```vb
Dim pr As Inventor.Parameter = ThisDoc.Document.componentdefinition.parameters("Offset2")
Dim v As Double = pr.Value/2.54
pr.Expression=1
pr.Expression =Round(v,1)
```

```vb
Dim pr As Inventor.Parameter = ThisDoc.Document.componentdefinition.parameters("Offset2")
Dim v As Double = pr.Value/2.54
pr.Expression=1
pr.Expression =Round(v,1)
```

```vb
oTrigger = Offset2
Dim pr As Inventor.Parameter = ThisDoc.Document.componentdefinition.parameters("Offset2")
Dim v As Double = pr.Value/2.54
pr.Expression=1
pr.Expression =Round(v,1)
```

---

## ControlDefinition.Execute2 does not run unless it is the last line in the rule

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361](https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361)

**Author:** Nick_Hall

**Date:** ‎10-05-2025
	
		
		09:06 PM

**Description:** I am writing some iLogic that interacts with the Vault Revision Table. One of the things I want to do as part of the rule is to update the Revision Table with the Vault data. I want to run it as an"After Open Document" rule.  This is a minimal version of the code I would like to run. MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox....

**Code:**

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox.Show("Updated Revision Table")
```

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox.Show("Updated Revision Table")
```

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)
```

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)
```

---

## Check if drawing parts list is splitted

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/check-if-drawing-parts-list-is-splitted/td-p/12568593](https://forums.autodesk.com/t5/inventor-programming-forum/check-if-drawing-parts-list-is-splitted/td-p/12568593)

**Author:** goran_nilssonK2TWB

**Date:** ‎02-19-2024
	
		
		12:06 AM

**Description:** Hello! I want to check if drawing parts list is splitted. If it's splitted I want to get a message.My code is using On error resume next, so it's not possible to use a Try.I can't find any information if it's possible or not. /Goran
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Public Sub Main()
	Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If oDDoc Is Nothing Then Exit Sub  
	Dim oBrowserPane As BrowserPane = oDDoc.BrowserPanes.ActivePane
	Dim oTopBrowserNode As BrowserNode = oBrowserPane.TopNode   
	Dim oBrowserNode As BrowserNode    
	For Each oBrowserNode In oTopBrowserNode.BrowserNodes
		Dim oBrowserNode2 As BrowserNode
		For Each oBrowserNode2 In oBrowserNode.BrowserNodes
			Dim oNativeObject As Object = oBrowserNode2.NativeObject
			If TypeName(oNativeObject) = "PartsList" Then
				If oBrowserNode2.BrowserNodes.Count > 0 Then
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split")
				End If
				Exit For
			End If
		Next
	Next       
End Sub
```

```vb
Public Sub Main()
	Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If oDDoc Is Nothing Then Exit Sub  
	Dim oBrowserPane As BrowserPane = oDDoc.BrowserPanes.ActivePane
	Dim oTopBrowserNode As BrowserNode = oBrowserPane.TopNode   
	Dim oBrowserNode As BrowserNode    
	For Each oBrowserNode In oTopBrowserNode.BrowserNodes
		Dim oBrowserNode2 As BrowserNode
		For Each oBrowserNode2 In oBrowserNode.BrowserNodes
			Dim oNativeObject As Object = oBrowserNode2.NativeObject
			If TypeName(oNativeObject) = "PartsList" Then
				If oBrowserNode2.BrowserNodes.Count > 0 Then
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split")
				End If
				Exit For
			End If
		Next
	Next       
End Sub
```

```vb
Sub main
	Dim DrawingDoc = ThisApplication.ActiveEditDocument
	Dim modelState As Integer = 0
	For Each oSheet In DrawingDoc.Sheets
		itemNumber=1
		While itemNumber < 10
			Try
				oPartsList = oSheet.PartsLists.Item(itemNumber)
				If modelState = oPartsList.ReferencedDocumentDescriptor.ReferencedModelState
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split or this is the first occurance of it")
					modelState=oPartsList.ReferencedDocumentDescriptor.ReferencedModelState
				End If
					itemNumber = itemNumber + 1
			Catch
				Exit While
			End Try
		End While
	Next
End Sub
```

---

## Load second vba project file or reference subroutine in other project file

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/load-second-vba-project-file-or-reference-subroutine-in-other/td-p/4861128#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/load-second-vba-project-file-or-reference-subroutine-in-other/td-p/4861128#messageview_0)

**Author:** pball

**Date:** ‎03-06-2014
	
		
		08:45 AM

**Description:** I'm looking to use two vba project files. I'll have my main ivb file and there is a second ivb file with a single script in it. I do not want to copy that script to the other project file. Is there a way to load or link the second project file from the main one?Another thing is I've manually loaded a second project file for testing and I can run scripts from it via the Macros option on the Tools ribbon bar but I cannot add them to the ribbon bar under the customize user commands option. If there...

**Code:**

```vb
Public Sub LoadVBAProject() LoadAnotherVBAProject "Put Location of VBA project here"End SubPublic Sub LoadAnotherVBAProject(pvarLocation As String) ' counter to work out the new project number Dim varVBACounter As Long varVBACounter = ThisApplication.VBAProjects.Count ' Open the other project ThisApplication.VBAProjects.Open pvarLocation ' get a reference to the just opened VBA project Dim varTempProject As InventorVBAProject Set varTempProject = ThisApplication.VBAProjects(varVBACounter + 1) ' references for all the different parts of the VBA project Dim varVBAComponent As InventorVBAComponent Dim varVBAMember As InventorVBAMember Dim varVBAArgument As InventorVBAArgument ' go through all the components of the Project and display them For Each varVBAComponent In varTempProject.InventorVBAComponents ' print to the immediate window the name of the component Debug.Print varVBAComponent.Name For Each varVBAMember In varVBAComponent.InventorVBAMembers ' print to the immediate window the name of the members of the component Debug.Print " " & varVBAMember.Name For Each varVBAArgument In varVBAMember.Arguments Debug.Print " " & varVBAArgument.Name & "----" & varVBAArgument.ArgumentType Next varVBAArgument Next varVBAMember Next varVBAComponentEnd Sub
```

---

## Code to measure length and width

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/code-to-measure-length-and-width/td-p/13792602#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/code-to-measure-length-and-width/td-p/13792602#messageview_0)

**Author:** berry.lejeune

**Date:** ‎09-02-2025
	
		
		03:23 AM

**Description:** Hi all, Don't even know if it's possible but I'll ask anywayWe've received a Revit model of a big building which is covered in stone plates. Now I need to get those dimensions of those plates for our manufacturer.But in stead of measuring each indivudial one, is it possible with a code to click on the surface and that it gives me the length and the width? Thanks

**Code:**

```vb
'Dimensions

Geo_Length = Measure.PreciseExtentsLength
Geo_Width = Measure.PreciseExtentsWidth
Geo_Height = Measure.PreciseExtentsHeight
```

```vb
Dim oInvApp As Inventor.Application = ThisApplication
Dim oCM As CommandManager = oInvApp.CommandManager
Dim oUOM As UnitsOfMeasure = oInvApp.ActiveDocument.UnitsOfMeasure
'Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kFootLengthUnits
Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kCentimeterLengthUnits

Do
	Dim oFace As Object
	oFace = oCM.Pick(SelectionFilterEnum.kPartFaceFilter, "Select Face...")
	If oFace Is Nothing Then Exit Do
	Dim oBox As Box2d = oFace.Evaluator.ParamRangeRect
	Dim oSizes(1) As Double
	oSizes(0) = Round(oUOM.ConvertUnits(oBox.MaxPoint.X - oBox.MinPoint.X,
										eLeng, oUOM.LengthUnits), 3)
	oSizes(1) = Round(oUOM.ConvertUnits(oBox.MaxPoint.Y - oBox.MinPoint.Y,
										eLeng, oUOM.LengthUnits), 3)
	MessageBox.Show("Length - " & Abs(oSizes.Max) & vbLf &
					"Width - " & Abs(oSizes.Min), "Surface size:")
	Logger.Info("Length - " & Abs(oSizes.Max))
	Logger.Info("Width - " & Abs(oSizes.Min))
Loop
```

```vb
Dim oInvApp As Inventor.Application = ThisApplication
Dim oCM As CommandManager = oInvApp.CommandManager
Dim oUOM As UnitsOfMeasure = oInvApp.ActiveDocument.UnitsOfMeasure
'Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kFootLengthUnits
Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kCentimeterLengthUnits

Do
	Dim oFace As Object
	oFace = oCM.Pick(SelectionFilterEnum.kPartFaceFilter, "Select Face...")
	If oFace Is Nothing Then Exit Do
	Dim oBox As Box2d = oFace.Evaluator.ParamRangeRect
	Dim oSizes(1) As Double
	oSizes(0) = Round(oUOM.ConvertUnits(oBox.MaxPoint.X - oBox.MinPoint.X,
										eLeng, oUOM.LengthUnits), 3)
	oSizes(1) = Round(oUOM.ConvertUnits(oBox.MaxPoint.Y - oBox.MinPoint.Y,
										eLeng, oUOM.LengthUnits), 3)
	MessageBox.Show("Length - " & Abs(oSizes.Max) & vbLf &
					"Width - " & Abs(oSizes.Min), "Surface size:")
	Logger.Info("Length - " & Abs(oSizes.Max))
	Logger.Info("Width - " & Abs(oSizes.Min))
Loop
```

---

## ActiveView.Fit does not update until after the rule finishes running.

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/activeview-fit-does-not-update-until-after-the-rule-finishes/td-p/13779026](https://forums.autodesk.com/t5/inventor-programming-forum/activeview-fit-does-not-update-until-after-the-rule-finishes/td-p/13779026)

**Author:** jsemansDCRCS

**Date:** ‎08-22-2025
	
		
		07:57 AM

**Description:** I am trying to generate a library of screenshots for a quick reference. The below code is what I am using to try and fit the model to the screen then grab the view and save it out. The problem is that .Fit does not seem to be updating the screen until the rule finishes running. The rule is set up to run on close, so sometimes an engineer will have closed the file while zoomed in causing the screenshot to be just the portion zoomed in on. ActiveView.GoHome works, but the home on some files is use...

**Code:**

```vb
Dim invApp As Inventor.Application = ThisApplication
invApp.ActiveView.Fit
invApp.ActiveView.SaveAsBitmap(savePath, 1150, 635)
```

```vb
Dim invApp As Inventor.Application = ThisApplication
invApp.ActiveView.Fit
invApp.ActiveView.SaveAsBitmap(savePath, 1150, 635)
```

---

## Export Drawing dimension to excel

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062#messageview_0)

**Author:** eladm

**Date:** ‎03-14-2021
	
		
		12:02 AM

**Description:** Hi I need help with ilogic rule to export all drawing dimensions to ExcelNot VBA , some of the dimeson  have toleranceregards
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	'MsgBox (b.Tolerance.Upper)
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	'MsgBox (b.Tolerance.Upper)
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	
	GoExcel.CellValue("A" & 1) = "Dimension"
	GoExcel.CellValue("B" & 1) = "Tolerance Upper"
	GoExcel.CellValue("C" & 1) = "Tolerance Lower"
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	
	GoExcel.CellValue("A" & 1) = "Dimension"
	GoExcel.CellValue("B" & 1) = "Tolerance Upper"
	GoExcel.CellValue("C" & 1) = "Tolerance Lower"
	GoExcel.Save

	Next
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper
oCells.item(i,3).value  = b.Tolerance.Lower
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper
oCells.item(i,3).value  = b.Tolerance.Lower
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet

Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet

Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add'.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

'oCells.item(2,1).value= b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'MsgBox (b.Text.Text & " " & b.Tolerance.ToleranceType)
	
	If b.Tolerance.ToleranceType = 31236 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31233 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31235 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = "+/-" & b.Tolerance.Upper*10
		'oCells.item(i,3).value  = "+/-" &b.Tolerance.Lower*10.ToString
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If


	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add'.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

'oCells.item(2,1).value= b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'MsgBox (b.Text.Text & " " & b.Tolerance.ToleranceType)
	
	If b.Tolerance.ToleranceType = 31236 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31233 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31235 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = "+/-" & b.Tolerance.Upper*10
		'oCells.item(i,3).value  = "+/-" &b.Tolerance.Lower*10.ToString
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If


	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

---

## Excluding Print and Count if Sheet name matches

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697](https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697)

**Author:** adeel.malik

**Date:** ‎09-01-2022
	
		
		08:03 PM

**Description:** I have a bunch of sheets named Vendor (e.g. Vendor:1 or Vendor:2...). For certain things I want to exclude these sheets from count and print. I wanted to write an ilogic for this but my code is not working.  Dim oDoc As DrawingDocument
oDoc = ThisApplication.ActiveDocument
Dim oSheet As Sheet

i = 0
x = 0

For Each oSheet In oDoc.Sheets
	i = i + 1
Next

While (x<i)
	Sheet_Name = "Vendor:" & x 
	x= x+1
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromPrinting = True  <---Not working
	'ThisDo...

**Code:**

```vb
Dim oDoc As DrawingDocument
oDoc = ThisApplication.ActiveDocument
Dim oSheet As Sheet

i = 0
x = 0

For Each oSheet In oDoc.Sheets
	i = i + 1
Next

While (x<i)
	Sheet_Name = "Vendor:" & x 
	x= x+1
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromPrinting = True  <---Not working
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromCount = True <---Not working
End While
```

```vb
Dim oDoc As DrawingDocument = ThisDoc.Document
Dim i as Integer = 0

For Each oSheet As Sheet In oDoc.Sheets
	
	i = i + 1
	
	If oSheet.Name = "Vendor:" & i
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
	End If
Next
```

```vb
Dim oDoc As DrawingDocument = ThisDoc.Document
Dim i as Integer = 0

For Each oSheet As Sheet In oDoc.Sheets
	
	i = i + 1
	
	If oSheet.Name = "Vendor:" & i
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
	End If
Next
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeVendorSheets As Boolean = InputRadioBox("Exclude Vendor Sheets?", "True", "False", False, "Toggle Vendor Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("Vendor") Then
		oSheet.ExcludeFromCount = oExcludeVendorSheets
		oSheet.ExcludeFromPrinting = oExcludeVendorSheets
	End If
Next
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeVendorSheets As Boolean = InputRadioBox("Exclude Vendor Sheets?", "True", "False", False, "Toggle Vendor Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("Vendor") Then
		oSheet.ExcludeFromCount = oExcludeVendorSheets
		oSheet.ExcludeFromPrinting = oExcludeVendorSheets
	End If
Next
```

```vb
If oSheet.Name.Contains("Vendor") Then
   oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
   oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
End If
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("S") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next
```

```vb
If oSheet.Name.StartsWith("S") Then
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.NotStartsWith("C") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If Not oSheet.Name.StartsWith("C") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next
```

---

## Custom ribbon button doesn't execute VBA sub

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-button-doesn-t-execute-vba-sub/td-p/13788549#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-button-doesn-t-execute-vba-sub/td-p/13788549#messageview_0)

**Author:** sv.lucht

**Date:** ‎08-29-2025
	
		
		04:54 AM

**Description:** I try to create custom ribbon tabs that get called when different criteria are met. The VBA code resides in StampMaker.ivb library in the UIManager module and everything works fine. I can create and destroy tabs,panel,buttons, but I cannot make the buttons work. The functions get never called. I tried to simplify the code, as seen below. I tried to get help from Gemini. But nothing works. So I would be very grateful for any suggestions. Here is a simplified version of the code just for troublesh...

**Code:**

```vb
Public Sub RunCleanButtonTest()
' --- Phase 1: Aggressive Cleanup ---
'Const TEST_MACRO_NAME As String = "Default.Module1.testio"
Const TEST_MACRO_NAME As String = "StampMaker.UIManager.testprompt"
Const TEST_TAB_NAME As String = "MyTestTab"
Const TEST_PANEL_NAME As String = "MyTestPanel"

' Delete the Control Definition
On Error Resume Next
ThisApplication.CommandManager.ControlDefinitions.item(TEST_MACRO_NAME).delete
Debug.Print "Old Control Definition deleted (if it existed)."
On Error GoTo 0

' Delete the Ribbon Tab
Dim zeroRibbon As Inventor.Ribbon
Set zeroRibbon = ThisApplication.UserInterfaceManager.Ribbons.item("ZeroDoc")
On Error Resume Next
zeroRibbon.RibbonTabs.item(TEST_TAB_NAME).delete
Debug.Print "Old Ribbon Tab deleted (if it existed)."
On Error GoTo 0

' --- Phase 2: Create a single, simple button ---
Debug.Print "Creating new UI..."

' Create Tab and Panel
Dim newTab As RibbonTab
Set newTab = zeroRibbon.RibbonTabs.Add("My Test", TEST_TAB_NAME, "ClientID_TestTab")
Dim newPanel As RibbonPanel
Set newPanel = newTab.RibbonPanels.Add("My Panel", TEST_PANEL_NAME, "ClientID_TestPanel")

' Create Button Definition
Dim buttonDef As ButtonDefinition
Set buttonDef = ThisApplication.CommandManager.ControlDefinitions.AddButtonDefinition( _
"Click Me", _
TEST_MACRO_NAME, _
CommandTypesEnum.kNonShapeEditCmdType)

' Add button to panel
Call newPanel.CommandControls.AddButton(buttonDef, True)

Debug.Print "--- Test UI created. Please click the 'Click Me' button on the 'My Test' tab. ---"
End Sub


Public Sub testprompt()
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
End Sub
```

```vb
Public Sub RunCleanButtonTest()
' --- Phase 1: Aggressive Cleanup ---
'Const TEST_MACRO_NAME As String = "Default.Module1.testio"
Const TEST_MACRO_NAME As String = "StampMaker.UIManager.testprompt"
Const TEST_TAB_NAME As String = "MyTestTab"
Const TEST_PANEL_NAME As String = "MyTestPanel"

' Delete the Control Definition
On Error Resume Next
ThisApplication.CommandManager.ControlDefinitions.item(TEST_MACRO_NAME).delete
Debug.Print "Old Control Definition deleted (if it existed)."
On Error GoTo 0

' Delete the Ribbon Tab
Dim zeroRibbon As Inventor.Ribbon
Set zeroRibbon = ThisApplication.UserInterfaceManager.Ribbons.item("ZeroDoc")
On Error Resume Next
zeroRibbon.RibbonTabs.item(TEST_TAB_NAME).delete
Debug.Print "Old Ribbon Tab deleted (if it existed)."
On Error GoTo 0

' --- Phase 2: Create a single, simple button ---
Debug.Print "Creating new UI..."

' Create Tab and Panel
Dim newTab As RibbonTab
Set newTab = zeroRibbon.RibbonTabs.Add("My Test", TEST_TAB_NAME, "ClientID_TestTab")
Dim newPanel As RibbonPanel
Set newPanel = newTab.RibbonPanels.Add("My Panel", TEST_PANEL_NAME, "ClientID_TestPanel")

' Create Button Definition
Dim buttonDef As ButtonDefinition
Set buttonDef = ThisApplication.CommandManager.ControlDefinitions.AddButtonDefinition( _
"Click Me", _
TEST_MACRO_NAME, _
CommandTypesEnum.kNonShapeEditCmdType)

' Add button to panel
Call newPanel.CommandControls.AddButton(buttonDef, True)

Debug.Print "--- Test UI created. Please click the 'Click Me' button on the 'My Test' tab. ---"
End Sub


Public Sub testprompt()
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
End Sub
```

```vb
Public Sub CreateMinimalMacroButton()
' --- 1. Define the names for the macro and UI elements ---
Const MACRO_FULL_NAME As String = "Module1.MyTargetSubroutine"
Const TAB_ID As String = "StampMaker_MinimalTab_1"
Const PANEL_ID As String = "StampMaker_MinimalPanel_1"

Dim uiMgr As UserInterfaceManager
Set uiMgr = ThisApplication.UserInterfaceManager

' --- 2. Clean up any old versions of these UI elements ---
' Using "On Error Resume Next" is a simple way to delete items
' without causing an error if they don't exist yet.
On Error Resume Next

' Delete the Ribbon Tab (this also removes the panel and the button on it)
uiMgr.Ribbons.item("ZeroDoc").RibbonTabs.item(TAB_ID).delete

' Delete the Control Definition (the "brain" of the button)
ThisApplication.CommandManager.ControlDefinitions.item(MACRO_FULL_NAME).delete

' Return to normal error handling
On Error GoTo 0

' --- 3. Create the new UI ---

' Get the ribbon that's visible when no document is open
Dim zeroRibbon As Ribbon
Set zeroRibbon = uiMgr.Ribbons.item("ZeroDoc")

' Create the new tab
Dim myTab As RibbonTab
Set myTab = zeroRibbon.RibbonTabs.Add("Minimal Test", TAB_ID, "ClientID_MinimalTab")

' Create a panel on the tab
Dim myPanel As RibbonPanel
Set myPanel = myTab.RibbonPanels.Add("Test Panel", PANEL_ID, "ClientID_MinimalPanel")

' Create the Macro Definition - this links the button to your code
Dim macroDef As MacroControlDefinition
Set macroDef = ThisApplication.CommandManager.ControlDefinitions.AddMacroControlDefinition(MACRO_FULL_NAME)

' Add the actual button to the panel
Call myPanel.CommandControls.AddMacro(macroDef)

' --- 4. Notify the user ---
MsgBox "Minimal macro button has been created." & vbCrLf & _
"Look for the 'Minimal Test' tab in your ribbon.", vbInformation

End Sub
```

```vb
Public Sub CreateMinimalMacroButton()
' --- 1. Define the names for the macro and UI elements ---
Const MACRO_FULL_NAME As String = "Module1.MyTargetSubroutine"
Const TAB_ID As String = "StampMaker_MinimalTab_1"
Const PANEL_ID As String = "StampMaker_MinimalPanel_1"

Dim uiMgr As UserInterfaceManager
Set uiMgr = ThisApplication.UserInterfaceManager

' --- 2. Clean up any old versions of these UI elements ---
' Using "On Error Resume Next" is a simple way to delete items
' without causing an error if they don't exist yet.
On Error Resume Next

' Delete the Ribbon Tab (this also removes the panel and the button on it)
uiMgr.Ribbons.item("ZeroDoc").RibbonTabs.item(TAB_ID).delete

' Delete the Control Definition (the "brain" of the button)
ThisApplication.CommandManager.ControlDefinitions.item(MACRO_FULL_NAME).delete

' Return to normal error handling
On Error GoTo 0

' --- 3. Create the new UI ---

' Get the ribbon that's visible when no document is open
Dim zeroRibbon As Ribbon
Set zeroRibbon = uiMgr.Ribbons.item("ZeroDoc")

' Create the new tab
Dim myTab As RibbonTab
Set myTab = zeroRibbon.RibbonTabs.Add("Minimal Test", TAB_ID, "ClientID_MinimalTab")

' Create a panel on the tab
Dim myPanel As RibbonPanel
Set myPanel = myTab.RibbonPanels.Add("Test Panel", PANEL_ID, "ClientID_MinimalPanel")

' Create the Macro Definition - this links the button to your code
Dim macroDef As MacroControlDefinition
Set macroDef = ThisApplication.CommandManager.ControlDefinitions.AddMacroControlDefinition(MACRO_FULL_NAME)

' Add the actual button to the panel
Call myPanel.CommandControls.AddMacro(macroDef)

' --- 4. Notify the user ---
MsgBox "Minimal macro button has been created." & vbCrLf & _
"Look for the 'Minimal Test' tab in your ribbon.", vbInformation

End Sub
```

---

## iLogic - Check active project is true

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-check-active-project-is-true/td-p/10218046](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-check-active-project-is-true/td-p/10218046)

**Author:** neil.cooke

**Date:** ‎04-07-2021
	
		
		02:19 AM

**Description:** I have found similar requests with solution but I am quite a new user to iLogic and perhaps missing the obvious (Sorry). I am trying to check that a specific project file is active, if it is not true then to pop up a message box stating you need to set the active project to the specified file. Please could someone help. Many thanks 
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Public Sub InventorProjectFile(oInvapp As Inventor.Application)

        Dim proj As DesignProject

        Try
            proj = oInvapp.DesignProjectManager.DesignProjects.ItemByName("C:\#######.ipj")
            proj.Activate()
        Catch ex As Exception
            If Not oInvapp.DesignProjectManager.ActiveDesignProject.FullFileName = "C:\#######.ipj" Then
                MsgBox("The wrong Project file has been Loaded !" & vbCrLf & "The Current Project file = " & oInvapp.DesignProjectManager.ActiveDesignProject.Name & vbCrLf & "Load the " & "C:\#######.ipj" & " file", vbCritical)
            End If
        End Try
end sub
```

```vb
Public Sub InventorProjectFile(oInvapp As Inventor.Application)

        Dim proj As DesignProject

        Try
            proj = oInvapp.DesignProjectManager.DesignProjects.ItemByName("C:\#######.ipj")
            proj.Activate()
        Catch ex As Exception
            If Not oInvapp.DesignProjectManager.ActiveDesignProject.FullFileName = "C:\#######.ipj" Then
                MsgBox("The wrong Project file has been Loaded !" & vbCrLf & "The Current Project file = " & oInvapp.DesignProjectManager.ActiveDesignProject.Name & vbCrLf & "Load the " & "C:\#######.ipj" & " file", vbCritical)
            End If
        End Try
end sub
```

```vb
Sub Main
	Const sProjPath As String = "C:\Temp\"
	Const sProjFileName As String = "MyInventorProject.ipj"
	Const sProjFile As String = "C:\Temp\MyInventorProject.ipj"
	Dim oProjMgr As Inventor.DesignProjectManager = ThisApplication.DesignProjectManager
	'get FullFileName of active DesignProject
	Dim sActiveProj As String = oProjMgr.ActiveDesignProject.FullFileName
	'only react if it is not the one expected / wanted
	If Not sActiveProj = sProjFile Then
		'Log it, or let user know
		Logger.Info("Active Project File Is:  " & sActiveProj)
		'MessageBox.Show("Active Project File Is:  " & sActiveProj)
		'try to find / get the DesignProject we want, by its FullFileName
		Dim oMyProj As Inventor.DesignProject = Nothing
		Try
			oMyProj = oProjMgr.DesignProjects.ItemByName(sProjFile)
		Catch
			Logger.Warn("Specified Inventor DesignProject Not Fount In Projects Collection!")
		End Try
		'if it was not found in that collection, then add it
		If oMyProj Is Nothing Then
			Try
				oMyProj = oProjMgr.DesignProjects.AddExisting(sProjFile)
			Catch
				Logger.Error("Error Adding Specified DesignProject To Projects Collection!")
			End Try
		End If
		'if we now have it, then activate it
		If oMyProj IsNot Nothing Then
			Try
				'True = Set as Default Project
				oMyProj.Activate(True)
			Catch
				Logger.Error("Error Activating Specified DesignProject!")
			End Try
		End If
	End If
End Sub
```

```vb
Sub Main
	Const sProjPath As String = "C:\Temp\"
	Const sProjFileName As String = "MyInventorProject.ipj"
	Const sProjFile As String = "C:\Temp\MyInventorProject.ipj"
	Dim oProjMgr As Inventor.DesignProjectManager = ThisApplication.DesignProjectManager
	'get FullFileName of active DesignProject
	Dim sActiveProj As String = oProjMgr.ActiveDesignProject.FullFileName
	'only react if it is not the one expected / wanted
	If Not sActiveProj = sProjFile Then
		'Log it, or let user know
		Logger.Info("Active Project File Is:  " & sActiveProj)
		'MessageBox.Show("Active Project File Is:  " & sActiveProj)
		'try to find / get the DesignProject we want, by its FullFileName
		Dim oMyProj As Inventor.DesignProject = Nothing
		Try
			oMyProj = oProjMgr.DesignProjects.ItemByName(sProjFile)
		Catch
			Logger.Warn("Specified Inventor DesignProject Not Fount In Projects Collection!")
		End Try
		'if it was not found in that collection, then add it
		If oMyProj Is Nothing Then
			Try
				oMyProj = oProjMgr.DesignProjects.AddExisting(sProjFile)
			Catch
				Logger.Error("Error Adding Specified DesignProject To Projects Collection!")
			End Try
		End If
		'if we now have it, then activate it
		If oMyProj IsNot Nothing Then
			Try
				'True = Set as Default Project
				oMyProj.Activate(True)
			Catch
				Logger.Error("Error Activating Specified DesignProject!")
			End Try
		End If
	End If
End Sub
```

---

## Section &amp; Detail View Label

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080](https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080)

**Author:** Talayoe

**Date:** ‎11-11-2021
	
		
		01:44 PM

**Description:** Good afternoon all, I am working with a large drawing package (for me at least) that is 36 pages long, 25 or so different parts for the build. What I dont care for is that the detail view labels are incrementing thru the entire drawing package and I would rathe have it sheet by sheet, if possible.  Example:Drawing Sheet 1: I create 3 section views, it auto labels then A-A, B-B, C-C, ect.Drawing Sheet 2: I create a single (or more) section views it currently labels them D-D, ect whereas I would l...

**Code:**

```vb
Dim oDoc As DrawingDocument
oDoc = ThisDoc.Document

Dim oSheet As Sheet
Dim oView As DrawingView
Dim oLabel As String
Dim RestartperSheet As Boolean

'Set this value to False if you want te rename views for all sheets
'Set this value to True if you want to rename views per sheet
RestartperSheet = True
'Set Start Label
oLabel = "A"

For Each oSheet In oDoc.Sheets
	'If True Set label Start Value per sheet
	If RestartperSheet = True
		oLabel = "A"
	End If
	For Each oView In oSheet.DrawingViews
		Logger.Info(oView.ViewType & " " & oView.Name)

		Select Case oView.ViewType
			Case kSectionDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kAuxiliaryDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kDetailDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kProjectedDrawingViewType
				If oView.Aligned = False Then
					'Set view name to Label
					oView.Name = oLabel
					'Set Label to next character
					oLabel = Chr(Asc(oLabel) + 1)
				End If
		End Select
	Next
Next
```

```vb
Dim oDoc As DrawingDocument
oDoc = ThisDoc.Document

Dim oStyleM As DrawingStylesManager
oStyleM = oDoc.StylesManager

Dim oStandard As DrawingStandardStyle
oStandard = oStyleM.ActiveStandardStyle

Dim oSheet As Sheet
Dim oView As DrawingView
Dim oLabel As String
Dim RestartperSheet As Boolean

'Set this value to False if you want te rename views for all sheets
'Set this value to True if you want to rename views per sheet
RestartperSheet = False
'Set Start Label
oLabel = "A"


For Each oSheet In oDoc.Sheets
	'If True Set label Start Value per sheet
	If RestartperSheet = True
		oLabel = "A"
	End If
	For Each oView In oSheet.DrawingViews
		Logger.Info(oView.ViewType & " " & oView.Name)
		'Exclude Charater from standard style list
		If oStandard.ApplyExcludeCharactersToViewNames = True Then
			While oStandard.ExcludeCharacters.Contains(oLabel) = True
				oLabel = Chr(Asc(oLabel) + 1)
			End While
		End If
		Select Case oView.ViewType
			Case kSectionDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kAuxiliaryDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kDetailDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kProjectedDrawingViewType
				If oView.Aligned = False Then
					'Set view name to Label
					oView.Name = oLabel
					'Set Label to next character
					oLabel = Chr(Asc(oLabel) + 1)
				End If
		End Select
	Next
Next

iLogicVb.UpdateWhenDone = True
```

```vb
oLabel = Chr(Asc(oLabel) + 1)
```

```vb
oLabel = GetAlphaString(oLabel)
```

```vb
Public Function GetAlphaString(ByVal ColumnLetter As String) As String
	'convert to number
	For i = 1 To Len(ColumnLetter)
		AlphaNum = AlphaNum * 26 + (Asc(UCase(Mid(ColumnLetter, i, 1))) -64)
	Next

	AlphaNum = AlphaNum + 1

	'convert back to format
	Do While AlphaNum > 0
		AlphaNum = AlphaNum - 1
		GetAlphaString = ChrW(65 + AlphaNum Mod 26) & GetAlphaString
		AlphaNum = AlphaNum \ 26
	Loop
End Function
```

```vb
Sub Main()
    Dim oDoc As DrawingDocument
    oDoc = ThisDoc.Document

    Dim oStyleM As DrawingStylesManager
    oStyleM = oDoc.StylesManager

    Dim oStandard As DrawingStandardStyle
    oStandard = oStyleM.ActiveStandardStyle

    Dim oSheet As Sheet
    Dim oView As DrawingView
    Dim oLabel As String
    Dim RestartperSheet As Boolean

    ' Set this value to True to rename views per sheet, False for continuous labeling across all sheets
    RestartperSheet = True
    ' Set Start Label
    oLabel = "A"

    For Each oSheet In oDoc.Sheets
        ' Reset label to A for each sheet if RestartperSheet is True
        If RestartperSheet = True Then
            oLabel = "A"
        End If
        For Each oView In oSheet.DrawingViews
            Logger.Info(oView.ViewType & " " & oView.Name)
            ' Skip excluded characters for the current label
            If oStandard.ApplyExcludeCharactersToViewNames = True Then
                While oStandard.ExcludeCharacters.Contains(oLabel) = True
                    oLabel = GetNextLabel(oLabel, oStandard)
                End While
            End If
            Select Case oView.ViewType
                Case kSectionDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kAuxiliaryDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kDetailDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kProjectedDrawingViewType
                    If oView.Aligned = False Then
                        ' Set view name to Label
                        oView.Name = oLabel
                        ' Get next valid label
                        oLabel = GetNextLabel(oLabel, oStandard)
                    End If
            End Select
        Next
    Next

    iLogicVb.UpdateWhenDone = True
End Sub

' Function to get the next valid label (A-Z, skipping excluded characters)
Function GetNextLabel(currentLabel As String, oStandard As DrawingStandardStyle) As String
    Dim nextLabel As String
    nextLabel = Chr(Asc(currentLabel) + 1)
    ' Ensure label stays within A-Z (ASCII 65-90)
    If Asc(nextLabel) > Asc("Z") Then
        nextLabel = "A" ' Reset to A if exceeding Z
    End If
    ' Skip excluded characters
    If oStandard.ApplyExcludeCharactersToViewNames = True Then
        While oStandard.ExcludeCharacters.Contains(nextLabel) = True
            nextLabel = Chr(Asc(nextLabel) + 1)
            If Asc(nextLabel) > Asc("Z") Then
                nextLabel = "A" ' Reset to A if exceeding Z
            End If
        End While
    End If
    GetNextLabel = nextLabel
End Function
```

---

## Inventor Vb.Net &quot;AddIn&quot; - Change A Drawing View

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/inventor-vb-net-quot-addin-quot-change-a-drawing-view/td-p/13828621#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/inventor-vb-net-quot-addin-quot-change-a-drawing-view/td-p/13828621#messageview_0)

**Author:** isocam

**Date:** ‎09-28-2025
	
		
		02:27 AM

**Description:** Can anybody help? I am creating a Vb.Net "AddIn" for Autodesk Inventor. I have a part (ipt) and a drawing (idw) document called.... 1000.ipt & 1000.idw I have created a copy the two documents so they become 2000.ipt & 2000.idw At this stage everything looks OK. However, When I open the drawing 2000.idw it still shows 1000.ipt for the drawing view(s). Does anybody know how I can, using Vb.Net, change the drawing view(s) so it references 2000.ipt and not 1000.ipt? I have tried "Copilot" and "ChatG...

**Code:**

```vb
Sub ReplaceModelReference(drawing As DrawingDocument, newFile As String)
    'This assumes the drawing references only one model

    ' EDIT: Line replaced as suggested by @WCrihfield 
    'Dim fileDescriptor As FileDescriptor = drawing.ReferencedFileDescriptors(1).DocumentDescriptor.ReferencedFileDescriptor
    Dim fileDescriptor As FileDescriptor = drawing.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor


    fileDescriptor.ReplaceReference(newFile)
End Sub
```

```vb
Sub ReplaceModelReference(drawing As DrawingDocument, newFile As String)
    'This assumes the drawing references only one model

    ' EDIT: Line replaced as suggested by @WCrihfield 
    'Dim fileDescriptor As FileDescriptor = drawing.ReferencedFileDescriptors(1).DocumentDescriptor.ReferencedFileDescriptor
    Dim fileDescriptor As FileDescriptor = drawing.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor


    fileDescriptor.ReplaceReference(newFile)
End Sub
```

---

## Switch active drawing Standard with iLogic

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/switch-active-drawing-standard-with-ilogic/td-p/13772691](https://forums.autodesk.com/t5/inventor-programming-forum/switch-active-drawing-standard-with-ilogic/td-p/13772691)

**Author:** Jesse_Glaze

**Date:** ‎08-18-2025
	
		
		11:27 AM

**Description:** I am looking for a (hopefully) simple code that would allow me to activate a different drawing standard.This rule would also be used for updating old drawings, so it should work even if the standard is saved to the style library but not saved in the local document.  
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Sub SetActiveDrawingStandardStyle(oDrawDoc As DrawingDocument, _
	Optional styleName As String = Nothing)
	Const sDefaultStyleName As String = "Default"
	If String.IsNullOrWhiteSpace(styleName) Then
		styleName = sDefaultStyleName
	End If
	Dim oSMgr As DrawingStylesManager = oDrawDoc.StylesManager
	Dim oStStyle As DrawingStandardStyle = Nothing
	Try
		oStStyle = oSMgr.StandardStyles.Item(styleName)
	Catch oEx As Exception
		Logger.Error("Failed to find specified DrawingStanardStyle to make active." _
		& vbCrLf & oEx.Message & vbCrLf & oEx.StackTrace)
	End Try
	If Not oStStyle Is Nothing Then
		If oStStyle.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
			oStStyle = oStStyle.ConvertToLocal()
		ElseIf oStStyle.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
			oStStyle.UpdateFromGlobal()
		End If
		oSMgr.ActiveStandardStyle = oStStyle
	End If
End Sub
```

```vb
Sub SetActiveDrawingStandardStyle(oDrawDoc As DrawingDocument, _
	Optional styleName As String = Nothing)
	Const sDefaultStyleName As String = "Default"
	If String.IsNullOrWhiteSpace(styleName) Then
		styleName = sDefaultStyleName
	End If
	Dim oSMgr As DrawingStylesManager = oDrawDoc.StylesManager
	Dim oStStyle As DrawingStandardStyle = Nothing
	Try
		oStStyle = oSMgr.StandardStyles.Item(styleName)
	Catch oEx As Exception
		Logger.Error("Failed to find specified DrawingStanardStyle to make active." _
		& vbCrLf & oEx.Message & vbCrLf & oEx.StackTrace)
	End Try
	If Not oStStyle Is Nothing Then
		If oStStyle.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
			oStStyle = oStStyle.ConvertToLocal()
		ElseIf oStStyle.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
			oStStyle.UpdateFromGlobal()
		End If
		oSMgr.ActiveStandardStyle = oStStyle
	End If
End Sub
```

---

## Rule with multiple options

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705#messageview_0)

**Author:** KaynePellegrino

**Date:** ‎08-26-2025
	
		
		09:28 AM

**Description:** This is a long one.------------------------ I make doors. They can be different Widths and Heights, among different Types.Widths:3-6 (41.75")3-0 (35.75")2-10 (33.75")2-8 (31.75")2-6 (29.75") Heights:8-0 (95.375")6-8 (79.375") Types:1K (1-Panel)2K (2-Panel3F (3-Panel)3S (3-Panel w/ Glass)50 (5-Panel) ------------------------ I have a Width Param & Rule:DoorWidthOptions [Text, Multi]: 3-6, 3-0, 2-10, 2-8, 2-6DoorWidth [Num]: (Varies)Door Width Rule:'Door Width
If DoorWidthOptions = "3-6" Then
		Do...

**Code:**

```vb
'Door Width
If DoorWidthOptions = "3-6" Then
		DoorWidth = 41.75
	End If

If DoorWidthOptions = "3-0" Then
		DoorWidth = 35.75
	End If

If DoorWidthOptions = "2-10" Then
		DoorWidth = 33.75
	End If

If DoorWidthOptions = "2-8" Then
		DoorWidth = 31.75
	End If

If DoorWidthOptions = "2-6" Then
		DoorWidth = 29.75
	End If
```

```vb
'Door Height
If DoorHeightOptions = "6-8" Then
		DoorHeight = 79.375
	End If

If DoorHeightOptions = "8-0" Then
		DoorHeight = 95.375
	End If
```

```vb
If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
		Component.IsActive("2K Skin: Back") = False
		Component.IsActive("2K Skin: Front") = False
		Component.IsActive("3F Skin: Back") = False
		Component.IsActive("3F Skin: Front") = False
		Component.IsActive("3F (3-6) Skin: Back") = False
		Component.IsActive("3F (3-6) Skin: Front") = False
		Component.IsActive("3S Skin: Back") = False
		Component.IsActive("3S Skin: Front") = False
		Component.IsActive("50 Skin: Back") = False
		Component.IsActive("50 Skin: Front") = False
	End If
```

```vb
If DoorType = "3F" And DoorWidthOptions = ("2-6" or "2-8" or "2-10" or "3-0") Then
		Component.IsActive("1K Skin: Back") = False
		Component.IsActive("1K Skin: Front") = False
		Component.IsActive("2K Skin: Back") = False
		Component.IsActive("2K Skin: Front") = False
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
		Component.IsActive("3F (3-6) Skin: Back") = False
		Component.IsActive("3F (3-6) Skin: Front") = False
		Component.IsActive("3S Skin: Back") = False
		Component.IsActive("3S Skin: Front") = False
		Component.IsActive("50 Skin: Back") = False
		Component.IsActive("50 Skin: Front") = False
	End If
```

```vb
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False
If DoorType = "3F" And DoorWidthO >= 29.75 And DoorWidthO <= 35.75
	Component.IsActive("3F Skin: Back") = True
	Component.IsActive("3F Skin: Front") = True
ElseIf DoorType = "3S" And DoorWidthO >= 29.75 And DoorWidthO <= 35.75
	Component.IsActive("3S Skin: Back") = True
	Component.IsActive("3S Skin: Front") = True
End If
```

```vb
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False
Select Case DoorType
Case "3F"
	Select Case DoorWidthO
	Case 29.75 To 35.75
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True	
	Case 2 To 29.75
		'Etc Etc Etc
	Case Else
		'Etc Etc Etc
	End Select
Case "3S"
	Select Case DoorWidthO
	Case 12 To 20
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case 2 To 12
		'Etc Etc Etc
	End Select
End Select
```

```vb
'"3F Skin" ActiveIf DoorType = "3F" And DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If
'"3F (3-6) Skin" Active
If DoorType = "3F" And DoorWidthOptions = "3-6" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
```

```vb
Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F"
	Select Case DoorWidthOptions
	Case 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 22
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 20
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 18
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select
```

```vb
Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F"
	Select Case DoorWidthOptions
	Case 18 to 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select
```

```vb
Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F", "3F Skin", "3F (3-6) Skin"
	Select Case DoorWidthOptions
	Case 18, 20, 22, 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	End If

If DoorType = "2K" Then
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If

If DoorType = "3S" Then
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	End If

If DoorType = "50" Then
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	End If
```

```vb
If DoorType = "3F" And DoorWidthOptions = "3-0" or "2-10" or "2-8" or "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End IfIf DoorType = "3F" And DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then                Component.IsActive("3F Skin: Back") = True                Component.IsActive("3F Skin: Front") = True        End If
```

```vb
'Door Height
If DoorHeightOptions = "6-8" Then
		DoorHeight = 79.375
	End If

If  DoorHeightOptions = "8-0" Then
		DoorHeight = 95.375
	End If


'Door Width
If DoorWidthOptions = "3-6" Then
		DoorWidth = 41.75
	End If

If DoorWidthOptions = "3-0" Then
		DoorWidth = 35.75
	End If

If DoorWidthOptions = "2-10" Then
		DoorWidth = 33.75
	End If

If DoorWidthOptions = "2-8" Then
		DoorWidth = 31.75
	End If

If DoorWidthOptions = "2-6" Then
		DoorWidth = 29.75
	End If
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	End If

If DoorType = "2K" Then
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth <= 40 Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth >= 40  And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If

If DoorType = "3S" Then
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	End If

If DoorType = "50" Then
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	End If
```

```vb
If DoorType = "3F" And DoorWidth <= 40 Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth >= 40  And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
```

```vb
If DoorType = "3F" Then
	If DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
ElseIf DoorType = "X" Then
	'Etc etc etc
ElseIf DoorType = "Y" Then
	'Etc etc etc
ElseIf DoorType = "Z" Then
	'Etc etc etc
End If
```

```vb
If DoorType = "3F" And DoorWidth = 35.75, 33.75, 31.75, AndOr 29.75 Then                "Component1" = True        ElseIf DoorWidth = 41.75 Then                "Component2" = True        End If
```

```vb
If "Variable1" = "OptionA1" And "Variable2" = "OptionB1" OR "OptionB2" Then '(Multiple Options in the Same Line, separated by some kind of 'or' statement)
		"Component1" = True                Else If "Variable2" = "OptionB3" Then                "Component2" = True
	End If
```

```vb
If DoorType = "3F" Then
	If DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
```

```vb
Sub main()
Dim X as integer = 1
Dim X1 as integer = 2 

Dim Y as integer = 3
Dim Y1 as integer = 4

If (x = x1 and y = y1) or y = 4 then
     'do something
Else if X = Y then 
     'Do something else
Else

End if 





End sub
```

```vb
Sub main()
Dim X as integer = 1
Dim X1 as integer = 2 

Dim Y as integer = 3
Dim Y1 as integer = 4

If (x = x1 and y = y1) or y = 4 then
     'do something
Else if X = Y then 
     'Do something else
Else

End if 





End sub
```

```vb
If (MiddleLetter.StartsWith("D") = True Or MiddleLetter.StartsWith("K") = True) And LastLetter.StartsWith("H") = False Then 
				Tank_Leg_Stand_TC = True	
			Else
				Tank_Leg_Stand_TC = False
			End If
```

```vb
If (MiddleLetter.StartsWith("D") = True Or MiddleLetter.StartsWith("K") = True) And LastLetter.StartsWith("H") = False Then 
				Tank_Leg_Stand_TC = True	
			Else
				Tank_Leg_Stand_TC = False
			End If
```

```vb
Dim A,B,C As String
A = "1"
B = "0"
C = "0"

If A = "1" And (B = "1" Or C = "1")
	MessageBox.Show("1", "Test")
Else
	MessageBox.Show("2", "Test")
end if
```

```vb
ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

Select Case DoorType
	Case "1K" 
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	Case "2K" 
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	Case "3F"
		If DoorWidthOptions <> "3-6"
			Component.IsActive("3F Skin: Back") = True
			Component.IsActive("3F Skin: Front") = True
		Else 
			If DoorHeightOptions = "8-0"
				Component.IsActive("3F (3-6) Skin: Back") = True
				Component.IsActive("3F (3-6) Skin: Front") = True
			Else
				MessageBox.Show(String.Format("DoorType = {0} | DoorWidthOptions = {1} | DoorHeightOptions = {2}", DoorType, DoorWidthOptions, DoorHeightOptions).ToString, "Un-Accounted For Result:")
			End If
		End If
	Case "3S" 
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case "50" 
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	Case Else
		MessageBox.Show(String.Format("DoorType = {0}", DoorType).ToString, "Un-Accounted For DoorType:")
End Select
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

Select Case DoorType
	Case "1K" 
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	Case "2K" 
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	Case "3F"
		If DoorWidthOptions <> "3-6"
			Component.IsActive("3F Skin: Back") = True
			Component.IsActive("3F Skin: Front") = True
		Else 
			If DoorHeightOptions = "8-0"
				Component.IsActive("3F (3-6) Skin: Back") = True
				Component.IsActive("3F (3-6) Skin: Front") = True
			Else
				MessageBox.Show(String.Format("DoorType = {0} | DoorWidthOptions = {1} | DoorHeightOptions = {2}", DoorType, DoorWidthOptions, DoorHeightOptions).ToString, "Un-Accounted For Result:")
			End If
		End If
	Case "3S" 
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case "50" 
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	Case Else
		MessageBox.Show(String.Format("DoorType = {0}", DoorType).ToString, "Un-Accounted For DoorType:")
End Select
```

```vb
If DoorWidthOptions <> "3-6"
```

```vb
Active_Condition = (DoorType = "3F" And (DoorWidthOptions = "2-6" Or DoorWidthOptions = "2-8" Or DoorWidthOptions = "2-10" Or DoorWidthOptions = "3-0"))

Component.IsActive("1K Skin: Back") = Active_Condition
Component.IsActive("1K Skin: Front") = Active_Condition
Component.IsActive("2K Skin: Back") = Not Active_Condition
Component.IsActive("2K Skin: Front") = Not Active_Condition
Component.IsActive("3F Skin: Back") = Not Active_Condition
Component.IsActive("3F Skin: Front") = Not Active_Condition
Component.IsActive("3F (3-6) Skin: Back") = Not Active_Condition
Component.IsActive("3F (3-6) Skin: Front") = Not Active_Condition
Component.IsActive("3S Skin: Back") = Not Active_Condition
Component.IsActive("3S Skin: Front") = Not Active_Condition
Component.IsActive("50 Skin: Back") = Not Active_Condition
Component.IsActive("50 Skin: Front") = Not Active_Condition
```

---

## Delete ImportedComponent from part environment

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/delete-importedcomponent-from-part-environment/td-p/12156512#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/delete-importedcomponent-from-part-environment/td-p/12156512#messageview_0)

**Author:** sneha.sadaphal

**Date:** ‎08-08-2023
	
		
		04:20 AM

**Description:** Hi,I have inserted the step file as imported component in inventor part environment.and also added one button to delete that imported component. but seems like the importedComponent.delete() is not implemented.getting error for delete();Is anyone have idea about this? if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition; ...

**Code:**

```vb
if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition;                
                ImportedComponents impComps = oPartCompdef.ReferenceComponents.ImportedComponents;
                foreach (ImportedComponent CompOcc in impComps)
                {                   
                  string fileUrn = CompOcc.Name.Replace(".stp","");
                   if (exchangeItem.ExchangeID.Contains(fileUrn))
                   {
                       CompOcc.Delete();
                   }
                }

            }
```

```vb
if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition;                
                ImportedComponents impComps = oPartCompdef.ReferenceComponents.ImportedComponents;
                foreach (ImportedComponent CompOcc in impComps)
                {                   
                  string fileUrn = CompOcc.Name.Replace(".stp","");
                   if (exchangeItem.ExchangeID.Contains(fileUrn))
                   {
                       CompOcc.Delete();
                   }
                }

            }
```

```vb
CompOcc.BreakLinkToFile();
```

```vb
CompOcc.BreakLinkToFile();
```

---

## Delete ImportedComponent from part environment

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/delete-importedcomponent-from-part-environment/td-p/12156512](https://forums.autodesk.com/t5/inventor-programming-forum/delete-importedcomponent-from-part-environment/td-p/12156512)

**Author:** sneha.sadaphal

**Date:** ‎08-08-2023
	
		
		04:20 AM

**Description:** Hi,I have inserted the step file as imported component in inventor part environment.and also added one button to delete that imported component. but seems like the importedComponent.delete() is not implemented.getting error for delete();Is anyone have idea about this? if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition; ...

**Code:**

```vb
if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition;                
                ImportedComponents impComps = oPartCompdef.ReferenceComponents.ImportedComponents;
                foreach (ImportedComponent CompOcc in impComps)
                {                   
                  string fileUrn = CompOcc.Name.Replace(".stp","");
                   if (exchangeItem.ExchangeID.Contains(fileUrn))
                   {
                       CompOcc.Delete();
                   }
                }

            }
```

```vb
if (_document != null && _document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument _oPart = _document as PartDocument;
                var oPartCompdef = _oPart.ComponentDefinition;                
                ImportedComponents impComps = oPartCompdef.ReferenceComponents.ImportedComponents;
                foreach (ImportedComponent CompOcc in impComps)
                {                   
                  string fileUrn = CompOcc.Name.Replace(".stp","");
                   if (exchangeItem.ExchangeID.Contains(fileUrn))
                   {
                       CompOcc.Delete();
                   }
                }

            }
```

```vb
CompOcc.BreakLinkToFile();
```

```vb
CompOcc.BreakLinkToFile();
```

---

## acad.exe link problem!

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/acad-exe-link-problem/td-p/13822188#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/acad-exe-link-problem/td-p/13822188#messageview_0)

**Author:** SergeLachance

**Date:** ‎09-23-2025
	
		
		07:43 AM

**Description:** i have a old rule who working great until a install inventor 2026??? i just change the new link but error message!any body can help me???sorry for my horrible english! my old rule: ' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveDocument
Dim oContext As T...

**Code:**

```vb
' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveDocument
Dim oContext As TranslationContext
oContext = ThisApplication.TransientObjects.CreateTranslationContext
oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
' Create a NameValueMap object
Dim oOptions As NameValueMap
oOptions = ThisApplication.TransientObjects.CreateNameValueMap
' Create a DataMedium object
Dim oDataMedium As DataMedium
oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
' Check whether the translator has 'SaveCopyAs' options
If DXFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DXFOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

DOSSIER = ThisDoc.Path
POSITION1 = InStrRev(DOSSIER, "\") 
POSITION2 = Right(DOSSIER, Len(DOSSIER) - POSITION1)

'get DXF target folder path
oFolder = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
    System.IO.Directory.CreateDirectory(oFolder)
End If

For i = 1 To 100
chiffre = i
myFile = oFolder & oFileName & "_Sheet_" & chiffre & ".dxf"
If(System.IO.File.Exists(myFile)) Then
Kill (myFile)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder & oFileName & ".dxf"
oFileName = ThisDoc.FileName(False) 'without extension

'Publish document.
DXFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder & oFileName & "_Sheet_" & 100 & ".dxf"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

' Get the DWG translator Add-In.
Dim oDWGAddIn As TranslatorAddIn 
oDWGAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
' Check whether the translator has 'SaveCopyAs' options
If oDWGAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DWGOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

'get DXF target folder path
oFolder2 = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder2) Then
    System.IO.Directory.CreateDirectory(oFolder2)
End If

For i = 1 To 100
chiffre2 = i
myFile2 = oFolder2 & oFileName2 & "_Sheet_" & chiffre2 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder2 & oFileName2 & ".dwg"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Publish document.
oDWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder2 & oFileName2 & "_Sheet_" & 100 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

question = MessageBox.Show("VEUX-TU OUVRIR AUTOCAD POUR PURGER TES DESSINS???", "OUVERTURE AUTOCAD???", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
MultiValue.SetValueOptions(True)
If question = vbYes Then
GoTo OUVERTURE
Else If question = vbNo Then
GoTo FIN
End If

OUVERTURE :
Dim acadExe = "C:\Program Files\Autodesk\AutoCAD 2025\acad.exe"

'MessageBox.Show("N'OUBLIE PAS DE SAUVER EN VERSION 2000, UNE FOIS TON DESSIN PURGER ET DE PURGER TOUTES TES PAGES", "ATTENTION",MessageBoxButtons.OK,MessageBoxIcon.Warning)
Process.Start(acadExe, oFolder2 & oFileName2 & "_Sheet_" & 1 & ".dwg")

FIN:
```

```vb
' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveDocument
Dim oContext As TranslationContext
oContext = ThisApplication.TransientObjects.CreateTranslationContext
oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
' Create a NameValueMap object
Dim oOptions As NameValueMap
oOptions = ThisApplication.TransientObjects.CreateNameValueMap
' Create a DataMedium object
Dim oDataMedium As DataMedium
oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
' Check whether the translator has 'SaveCopyAs' options
If DXFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DXFOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

DOSSIER = ThisDoc.Path
POSITION1 = InStrRev(DOSSIER, "\") 
POSITION2 = Right(DOSSIER, Len(DOSSIER) - POSITION1)

'get DXF target folder path
oFolder = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
    System.IO.Directory.CreateDirectory(oFolder)
End If

For i = 1 To 100
chiffre = i
myFile = oFolder & oFileName & "_Sheet_" & chiffre & ".dxf"
If(System.IO.File.Exists(myFile)) Then
Kill (myFile)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder & oFileName & ".dxf"
oFileName = ThisDoc.FileName(False) 'without extension

'Publish document.
DXFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder & oFileName & "_Sheet_" & 100 & ".dxf"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

' Get the DWG translator Add-In.
Dim oDWGAddIn As TranslatorAddIn 
oDWGAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
' Check whether the translator has 'SaveCopyAs' options
If oDWGAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DWGOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

'get DXF target folder path
oFolder2 = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder2) Then
    System.IO.Directory.CreateDirectory(oFolder2)
End If

For i = 1 To 100
chiffre2 = i
myFile2 = oFolder2 & oFileName2 & "_Sheet_" & chiffre2 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder2 & oFileName2 & ".dwg"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Publish document.
oDWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder2 & oFileName2 & "_Sheet_" & 100 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

question = MessageBox.Show("VEUX-TU OUVRIR AUTOCAD POUR PURGER TES DESSINS???", "OUVERTURE AUTOCAD???", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
MultiValue.SetValueOptions(True)
If question = vbYes Then
GoTo OUVERTURE
Else If question = vbNo Then
GoTo FIN
End If

OUVERTURE :
Dim acadExe = "‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"

'MessageBox.Show("N'OUBLIE PAS DE SAUVER EN VERSION 2000, UNE FOIS TON DESSIN PURGER ET DE PURGER TOUTES TES PAGES", "ATTENTION",MessageBoxButtons.OK,MessageBoxIcon.Warning)
Process.Start(acadExe, oFolder2 & oFileName2 & "_Sheet_" & 1 & ".dwg")

FIN:
```

```vb
Dim acadExe = "‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"

'MessageBox.Show("N'OUBLIE PAS DE SAUVER EN VERSION 2000, UNE FOIS TON DESSIN PURGER ET DE PURGER TOUTES TES PAGES", "ATTENTION",MessageBoxButtons.OK,MessageBoxIcon.Warning)
Process.Start(acadExe)
```

```vb
"‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"
```

```vb
Dim acadExe = "‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"

Process.Start(acadExe)
```

```vb
Process.Start(acad.exe)
```

```vb
C:\Program Files\Autodesk\AutoCAD 2026\acad.exe
```

```vb
Process.Start("notepad.exe")
```

---

## Add leading zeros to title block

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111](https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111)

**Author:** ludmylla.camielli

**Date:** ‎08-31-2020
	
		
		06:06 PM

**Description:** Hi guys, I need to fill my drawing's title block with the following, which will be the document number: XXXX.XXX.XXX.<Number of Sheets><Sheet Number> However, both numbers must have a leading zero (for example 0501, meaning 5 sheets and I'm currently looking at sheet 1). I have spent a couple of hours looking for any solution to this and ended making the following iLogic code, which I expected to be useful somehow. But I have no idea how to sync this with my document number. Is this the best way...

**Code:**

```vb
Dim oSheet As Sheet
Dim SheetNumber As Double

For Each oSheet In ThisApplication.ActiveDocument.Sheets
SheetNumber  = Mid(oSheet.Name, InStr(1, oSheet.Name, ":") + 1)
MessageBox.Show(Microsoft.VisualBasic.Strings.Format(SheetNumber,"0#"), oSheet.Name)
Next
```

```vb
'Get Total Number of Sheets
SheetCount = ThisApplication.ActiveDocument.Sheets.Count

'Append a leading zero if less than 10
If SheetCount < 10 Then
	SheetCo = "0" & SheetCount
Else 
	SheetCo = SheetCount
End If

'Loop through each sheet and grab the sheet number from the name
For Each Sheet In ThisApplication.ActiveDocument.Sheets
SheetNumber = Mid(Sheet.Name, InStr(1, Sheet.Name, ":") + 1)

'Append a leading zero if less than 10
If SheetNumber <10 Then 
SheetNo = "0" & SheetNumber
Else
	SheetNo = SheetNumber
End If

MessageBox.Show(SheetCo & SheetNo, "Sheet descriptor")

Next
```

```vb
'Get Total Number of Sheets
SheetCount = ThisApplication.ActiveDocument.Sheets.Count

'Append a leading zero if less than 10
If SheetCount < 10 Then
	SheetCo = "0" & SheetCount
Else 
	SheetCo = SheetCount
End If

'Loop through each sheet and grab the sheet number from the name
For Each Sheet In ThisApplication.ActiveDocument.Sheets
SheetNumber = Mid(Sheet.Name, InStr(1, Sheet.Name, ":") + 1)

'Append a leading zero if less than 10
If SheetNumber <10 Then 
SheetNo = "0" & SheetNumber
Else
	SheetNo = SheetNumber
End If

MessageBox.Show(SheetCo & SheetNo, "Sheet descriptor")

Next
```

```vb
Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)).ToString("00")
	MsgBox(oSheets.Count.ToString("00") & oSheetNumber)
Next
```

```vb
Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)).ToString("00")
	Dim oTB As TitleBlock = oSheet.TitleBlock
	For Each oTextBox As Inventor.TextBox In oTB.Definition.Sketch.TextBoxes
		If oTextBox.Text = "<CustomSheetNumber>"
			oTB.SetPromptResultText(oTextBox, (oSheets.Count.ToString("00") & oSheetNumber))
			Exit For
		End If
	Next
Next
```

```vb
Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)+SheetStartNum-1).ToString("0000")
	
	Dim oTB As TitleBlock = oSheet.TitleBlock
	For Each oTextBox As Inventor.TextBox In oTB.Definition.Sketch.TextBoxes
		If oTextBox.Text = "<CustomSheetNumber>"
			oTB.SetPromptResultText(oTextBox, (oSheetNumber))
			Exit For
		End If
	Next
Next
```

---

## ActiveView.Fit does not update until after the rule finishes running.

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/activeview-fit-does-not-update-until-after-the-rule-finishes/td-p/13779026#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/activeview-fit-does-not-update-until-after-the-rule-finishes/td-p/13779026#messageview_0)

**Author:** jsemansDCRCS

**Date:** ‎08-22-2025
	
		
		07:57 AM

**Description:** I am trying to generate a library of screenshots for a quick reference. The below code is what I am using to try and fit the model to the screen then grab the view and save it out. The problem is that .Fit does not seem to be updating the screen until the rule finishes running. The rule is set up to run on close, so sometimes an engineer will have closed the file while zoomed in causing the screenshot to be just the portion zoomed in on. ActiveView.GoHome works, but the home on some files is use...

**Code:**

```vb
Dim invApp As Inventor.Application = ThisApplication
invApp.ActiveView.Fit
invApp.ActiveView.SaveAsBitmap(savePath, 1150, 635)
```

```vb
Dim invApp As Inventor.Application = ThisApplication
invApp.ActiveView.Fit
invApp.ActiveView.SaveAsBitmap(savePath, 1150, 635)
```

---

## Section &amp; Detail View Label

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080#messageview_0)

**Author:** Talayoe

**Date:** ‎11-11-2021
	
		
		01:44 PM

**Description:** Good afternoon all, I am working with a large drawing package (for me at least) that is 36 pages long, 25 or so different parts for the build. What I dont care for is that the detail view labels are incrementing thru the entire drawing package and I would rathe have it sheet by sheet, if possible.  Example:Drawing Sheet 1: I create 3 section views, it auto labels then A-A, B-B, C-C, ect.Drawing Sheet 2: I create a single (or more) section views it currently labels them D-D, ect whereas I would l...

**Code:**

```vb
Dim oDoc As DrawingDocument
oDoc = ThisDoc.Document

Dim oSheet As Sheet
Dim oView As DrawingView
Dim oLabel As String
Dim RestartperSheet As Boolean

'Set this value to False if you want te rename views for all sheets
'Set this value to True if you want to rename views per sheet
RestartperSheet = True
'Set Start Label
oLabel = "A"

For Each oSheet In oDoc.Sheets
	'If True Set label Start Value per sheet
	If RestartperSheet = True
		oLabel = "A"
	End If
	For Each oView In oSheet.DrawingViews
		Logger.Info(oView.ViewType & " " & oView.Name)

		Select Case oView.ViewType
			Case kSectionDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kAuxiliaryDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kDetailDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kProjectedDrawingViewType
				If oView.Aligned = False Then
					'Set view name to Label
					oView.Name = oLabel
					'Set Label to next character
					oLabel = Chr(Asc(oLabel) + 1)
				End If
		End Select
	Next
Next
```

```vb
Dim oDoc As DrawingDocument
oDoc = ThisDoc.Document

Dim oStyleM As DrawingStylesManager
oStyleM = oDoc.StylesManager

Dim oStandard As DrawingStandardStyle
oStandard = oStyleM.ActiveStandardStyle

Dim oSheet As Sheet
Dim oView As DrawingView
Dim oLabel As String
Dim RestartperSheet As Boolean

'Set this value to False if you want te rename views for all sheets
'Set this value to True if you want to rename views per sheet
RestartperSheet = False
'Set Start Label
oLabel = "A"


For Each oSheet In oDoc.Sheets
	'If True Set label Start Value per sheet
	If RestartperSheet = True
		oLabel = "A"
	End If
	For Each oView In oSheet.DrawingViews
		Logger.Info(oView.ViewType & " " & oView.Name)
		'Exclude Charater from standard style list
		If oStandard.ApplyExcludeCharactersToViewNames = True Then
			While oStandard.ExcludeCharacters.Contains(oLabel) = True
				oLabel = Chr(Asc(oLabel) + 1)
			End While
		End If
		Select Case oView.ViewType
			Case kSectionDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kAuxiliaryDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kDetailDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kProjectedDrawingViewType
				If oView.Aligned = False Then
					'Set view name to Label
					oView.Name = oLabel
					'Set Label to next character
					oLabel = Chr(Asc(oLabel) + 1)
				End If
		End Select
	Next
Next

iLogicVb.UpdateWhenDone = True
```

```vb
oLabel = Chr(Asc(oLabel) + 1)
```

```vb
oLabel = GetAlphaString(oLabel)
```

```vb
Public Function GetAlphaString(ByVal ColumnLetter As String) As String
	'convert to number
	For i = 1 To Len(ColumnLetter)
		AlphaNum = AlphaNum * 26 + (Asc(UCase(Mid(ColumnLetter, i, 1))) -64)
	Next

	AlphaNum = AlphaNum + 1

	'convert back to format
	Do While AlphaNum > 0
		AlphaNum = AlphaNum - 1
		GetAlphaString = ChrW(65 + AlphaNum Mod 26) & GetAlphaString
		AlphaNum = AlphaNum \ 26
	Loop
End Function
```

```vb
Sub Main()
    Dim oDoc As DrawingDocument
    oDoc = ThisDoc.Document

    Dim oStyleM As DrawingStylesManager
    oStyleM = oDoc.StylesManager

    Dim oStandard As DrawingStandardStyle
    oStandard = oStyleM.ActiveStandardStyle

    Dim oSheet As Sheet
    Dim oView As DrawingView
    Dim oLabel As String
    Dim RestartperSheet As Boolean

    ' Set this value to True to rename views per sheet, False for continuous labeling across all sheets
    RestartperSheet = True
    ' Set Start Label
    oLabel = "A"

    For Each oSheet In oDoc.Sheets
        ' Reset label to A for each sheet if RestartperSheet is True
        If RestartperSheet = True Then
            oLabel = "A"
        End If
        For Each oView In oSheet.DrawingViews
            Logger.Info(oView.ViewType & " " & oView.Name)
            ' Skip excluded characters for the current label
            If oStandard.ApplyExcludeCharactersToViewNames = True Then
                While oStandard.ExcludeCharacters.Contains(oLabel) = True
                    oLabel = GetNextLabel(oLabel, oStandard)
                End While
            End If
            Select Case oView.ViewType
                Case kSectionDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kAuxiliaryDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kDetailDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kProjectedDrawingViewType
                    If oView.Aligned = False Then
                        ' Set view name to Label
                        oView.Name = oLabel
                        ' Get next valid label
                        oLabel = GetNextLabel(oLabel, oStandard)
                    End If
            End Select
        Next
    Next

    iLogicVb.UpdateWhenDone = True
End Sub

' Function to get the next valid label (A-Z, skipping excluded characters)
Function GetNextLabel(currentLabel As String, oStandard As DrawingStandardStyle) As String
    Dim nextLabel As String
    nextLabel = Chr(Asc(currentLabel) + 1)
    ' Ensure label stays within A-Z (ASCII 65-90)
    If Asc(nextLabel) > Asc("Z") Then
        nextLabel = "A" ' Reset to A if exceeding Z
    End If
    ' Skip excluded characters
    If oStandard.ApplyExcludeCharactersToViewNames = True Then
        While oStandard.ExcludeCharacters.Contains(nextLabel) = True
            nextLabel = Chr(Asc(nextLabel) + 1)
            If Asc(nextLabel) > Asc("Z") Then
                nextLabel = "A" ' Reset to A if exceeding Z
            End If
        End While
    End If
    GetNextLabel = nextLabel
End Function
```

---

## Reset view lables

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/reset-view-lables/td-p/10468556#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/reset-view-lables/td-p/10468556#messageview_0)

**Author:** sschulteH6WZ3

**Date:** ‎07-14-2021
	
		
		08:09 AM

**Description:** Hi all, have the programmers of Inventor seen their way to allow view letter \ number reset with out going through a bunch of time wasting gymnastics? I have seen people complaining about this on here since 08.

**Code:**

```vb
Sub Main()

Logger.Info(" ")
Logger.Info("RuleName: " & iLogicVb.RuleName)

'Re-numbers all view labels on all sheets
'	handles aplha or numeric
'	option to set all to aplha
'	option to set all to Numeric
'	option to set only details and sections to alpha


oScheme = "Numeric w/ details and sections as Alpha"
'oScheme = "All Numeric"
'oScheme = "All Alpha"

Dim oDoc As DrawingDocument
Dim oSheet As Sheet
Dim oView As DrawingView

'check if active document is a drawing
If ThisApplication.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
	Exit Sub
End If

oDoc = ThisApplication.ActiveDocument

i = 9
k = 1
'Rename view identifiers 
'look at all sheets
For Each oSheet In oDoc.Sheets
	'look at all views on sheet
	For Each oView In oSheet.DrawingViews

		'renumber only detail and section
		If oScheme = "All Numeric" Then
			oView.Name = i
			i = i + 1
		ElseIf oScheme = "All Alpha" Then
			oView.Name = num2Letter(i)
			i = i + 1

		Else If oScheme = "Numeric w/ details and sections as Alpha" Then

			If oView.ViewType = DrawingViewTypeEnum.kSectionDrawingViewType Or _
				oView.ViewType = DrawingViewTypeEnum.kDetailDrawingViewType Then
				oView.Name = num2Letter(i)
				
				'skip I, O, & Q
				If InStr(1, num2Letter(i), "I") > 0 Or _
					InStr(1, num2Letter(i), "O") > 0 Or _
					InStr(1, num2Letter(i), "Q") > 0 Then
					i = i + 1
					oView.Name = num2Letter(i)
				End If
				
				i = i + 1
			Else
				oView.Name = k
				k = k + 1
			End If
		End If
	Next oView
Next oSheet

End Sub

Public Function num2Letter(num As Long) As String
	'converts long to corresponding alpha
	'Ex. 1 = A, 28 = AB

	remain = num Mod 26
	whole = Fix(num / 26)

	If num < 27 Then
		If remain = 0 Then
			num2Letter = "Z"
		Else
			num2Letter = Chr(remain + 64)
		End If
	Else
		If remain = 0 Then
			num2Letter = Chr(whole + 63) & "Z"
		Else
			num2Letter = Chr(whole + 64) & Chr(remain + 64)
		End If
	End If
End Function
```

---

## Trailing Zeros

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427](https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427)

**Author:** cgchad

**Date:** ‎09-09-2025
	
		
		10:09 AM

**Description:** I know this has been asked dozens of times but for some reason I can't get it to work out.  I am running Inventor Professional 2024.I have 2 parameters that i need to display with just 2 decimal places.  Even if those decimal places are zeros.  No matter what I have done it stops when it's a whole number or has something like .70 it will change to .7. My code for this is currently:'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)
'set ra...

**Code:**

```vb
'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)
'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & WIDTH & " x " & LENGTH
```

```vb
Dim Height As Double = tank_id
Dim Width As Double = tank_height

Dim HeightString As String = CStr(Height)
Dim WidthString As String = CStr(Width)

Dim HeightParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_id")
Dim WidthParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_height")

HeightParam.Precision = 3 'I wish this worked to control the decimals displayed from either the expression or value. 
WidthParam.Precision = 3

Dim HeightSubString As String() = HeightParam.Expression.Split(" "c) 
Dim WidthSubString As String() = WidthParam.Expression.Split(" "c)


Logger.Info("Height blue value: " & tank_id & " Height double: " & Height & " Height String: " & HeightString & " Height param expression: " & HeightParam.Expression & " Height param value direct: " & HeightParam.Value / 2.54 & " Height substring: " & HeightSubString(0))
Logger.Info("Width blue value: "  & tank_height & " Widdth double: " & Width & "  Width String: " & WidthString & " Width param expression: " & WidthParam.Expression & " Width param value direct: " & HeightParam.Value / 2.54 & " Width substring: " & WidthSubString(0))
```

```vb
Dim Height As Double = tank_id
Dim Width As Double = tank_height

Dim HeightString As String = CStr(Height)
Dim WidthString As String = CStr(Width)

Dim HeightParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_id")
Dim WidthParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_height")

HeightParam.Precision = 3 'I wish this worked to control the decimals displayed from either the expression or value. 
WidthParam.Precision = 3

Dim HeightSubString As String() = HeightParam.Expression.Split(" "c) 
Dim WidthSubString As String() = WidthParam.Expression.Split(" "c)


Logger.Info("Height blue value: " & tank_id & " Height double: " & Height & " Height String: " & HeightString & " Height param expression: " & HeightParam.Expression & " Height param value direct: " & HeightParam.Value / 2.54 & " Height substring: " & HeightSubString(0))
Logger.Info("Width blue value: "  & tank_height & " Widdth double: " & Width & "  Width String: " & WidthString & " Width param expression: " & WidthParam.Expression & " Width param value direct: " & HeightParam.Value / 2.54 & " Width substring: " & WidthSubString(0))
```

```vb
'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
Dim DoubleLength As Double = Round(SheetMetal.FlatExtentsLength, 2)
Dim DoubleWidth As Double = Round(SheetMetal.FlatExtentsWidth, 2)
LENGTH = DoubleLength
WIDTH = DoubleWidth

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & DoubleWidth.ToString("F2") & " x " & DoubleLength.ToString("F2")
```

```vb
Dim DoubleLength As Double = Round(SheetMetal.FlatExtentsLength, 2)
Dim DoubleWidth As Double = Round(SheetMetal.FlatExtentsWidth, 2)
LENGTH = DoubleLength
WIDTH = DoubleWidth

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & DoubleWidth.ToString("F2") & " x " & DoubleLength.ToString("F2")
```

---

## Switch active drawing Standard with iLogic

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/switch-active-drawing-standard-with-ilogic/td-p/13772691#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/switch-active-drawing-standard-with-ilogic/td-p/13772691#messageview_0)

**Author:** Jesse_Glaze

**Date:** ‎08-18-2025
	
		
		11:27 AM

**Description:** I am looking for a (hopefully) simple code that would allow me to activate a different drawing standard.This rule would also be used for updating old drawings, so it should work even if the standard is saved to the style library but not saved in the local document.  
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Sub SetActiveDrawingStandardStyle(oDrawDoc As DrawingDocument, _
	Optional styleName As String = Nothing)
	Const sDefaultStyleName As String = "Default"
	If String.IsNullOrWhiteSpace(styleName) Then
		styleName = sDefaultStyleName
	End If
	Dim oSMgr As DrawingStylesManager = oDrawDoc.StylesManager
	Dim oStStyle As DrawingStandardStyle = Nothing
	Try
		oStStyle = oSMgr.StandardStyles.Item(styleName)
	Catch oEx As Exception
		Logger.Error("Failed to find specified DrawingStanardStyle to make active." _
		& vbCrLf & oEx.Message & vbCrLf & oEx.StackTrace)
	End Try
	If Not oStStyle Is Nothing Then
		If oStStyle.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
			oStStyle = oStStyle.ConvertToLocal()
		ElseIf oStStyle.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
			oStStyle.UpdateFromGlobal()
		End If
		oSMgr.ActiveStandardStyle = oStStyle
	End If
End Sub
```

```vb
Sub SetActiveDrawingStandardStyle(oDrawDoc As DrawingDocument, _
	Optional styleName As String = Nothing)
	Const sDefaultStyleName As String = "Default"
	If String.IsNullOrWhiteSpace(styleName) Then
		styleName = sDefaultStyleName
	End If
	Dim oSMgr As DrawingStylesManager = oDrawDoc.StylesManager
	Dim oStStyle As DrawingStandardStyle = Nothing
	Try
		oStStyle = oSMgr.StandardStyles.Item(styleName)
	Catch oEx As Exception
		Logger.Error("Failed to find specified DrawingStanardStyle to make active." _
		& vbCrLf & oEx.Message & vbCrLf & oEx.StackTrace)
	End Try
	If Not oStStyle Is Nothing Then
		If oStStyle.StyleLocation = StyleLocationEnum.kLibraryStyleLocation Then
			oStStyle = oStStyle.ConvertToLocal()
		ElseIf oStStyle.StyleLocation = StyleLocationEnum.kBothStyleLocation Then
			oStStyle.UpdateFromGlobal()
		End If
		oSMgr.ActiveStandardStyle = oStStyle
	End If
End Sub
```

---

## How to determine a simplified part via iLogic

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/how-to-determine-a-simplified-part-via-ilogic/td-p/13750575](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-determine-a-simplified-part-via-ilogic/td-p/13750575)

**Author:** dnguyennn

**Date:** ‎08-01-2025
	
		
		03:12 PM

**Description:** Is there a way to better determine a part is a simplified part in iLogic and VBA?except the string "Simplify" in the name and the activedoc is a PartDocumentI also found get ShrinkwrapComponents count is also kinda work?If oDef.ReferenceComponents.ShrinkwrapComponents.Count > 0 Then
    isSimplified = True
End If

**Code:**

```vb
If oDef.ReferenceComponents.ShrinkwrapComponents.Count > 0 Then
    isSimplified = True
End If
```

```vb
If oDef.ReferenceComponents.ShrinkwrapComponents.Count > 0 Then
    isSimplified = True
End If
```

---

## How to refresh/open the projects window VBA

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/how-to-refresh-open-the-projects-window-vba/td-p/13725927](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-refresh-open-the-projects-window-vba/td-p/13725927)

**Author:** gablewatson44

**Date:** ‎07-15-2025
	
		
		07:37 AM

**Description:** Currently using inventor 2024, i have written a macro that automatically creates a workgroup path based on the file location of whatever is opened. I am using a work around with Sendkeys that quickly opens and closes the project window using a custom shortcut:' Type P, R, O then EscapeSendKeys "P", TrueSendKeys "R", TrueSendKeys "O", TrueSendKeys "{ESC}", True while this works, i would rather use a form to completely control this macro thus requiring a different method of refreshing the project ...

**Code:**

```vb
Dim oControlDef As ControlDefinition
        
Set oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppProjectsCmd")
        
Call oControlDef.Execute
```

```vb
Dim oControlDef As ControlDefinition
        
Set oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppProjectsCmd")
        
Call oControlDef.Execute
```

---

## Inventor DWG to AutoCAD DWG Conversion

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/inventor-dwg-to-autocad-dwg-conversion/td-p/11339593](https://forums.autodesk.com/t5/inventor-programming-forum/inventor-dwg-to-autocad-dwg-conversion/td-p/11339593)

**Author:** Harshil.JAINS8RVT

**Date:** ‎08-04-2022
	
		
		02:29 AM

**Description:** Hello, I have a requirement to convert from Inventor DWG to AutoCAD DWG. I have tried the code as shown below:- ApplicationAddInsPtr pAddIns = pApplication->GetApplicationAddIns();TranslatorAddInPtr pTrans;_bstr_t ClientId = _T("{C24E3AC2-122E-11D5-8E91-0010B541CD80}");pTrans = pAddIns->GetItemById(ClientId);DocumentPtr pINVDoc = pApplication->GetActiveDocument();short bHasSaveOption = pTrans->GetHasSaveCopyAsOptions(pINVDoc, oContext, oOptions);if (bHasSaveOption){_bstr_t bstrAttribValue("D:\\C...

**Code:**

```vb
class ThisRule { 
public:
    CComPtr<Application> ThisApplication;

    void Main() {

        HRESULT Result = NOERROR;

        CComPtr<Document> gDoc;
        Result = ThisApplication->get_ActiveDocument(&gDoc);

        CComPtr<DrawingDocument> doc;
        doc = gDoc;
        
        CComPtr<TranslationContext> context;
        Result = ThisApplication->TransientObjects->CreateTranslationContext(&context);
        context->put_Type(kFileBrowseIOMechanism);

        CComPtr<NameValueMap> options;
        Result = ThisApplication->TransientObjects->CreateNameValueMap(&options);

        CComPtr<DataMedium> dataMedium;
        Result = ThisApplication->TransientObjects->CreateDataMedium(&dataMedium);

        CComPtr<TranslatorAddIn> translator;
        translator = GetDwgTranslator();


        if (translator->HasSaveCopyAsOptions[doc][context][options]) {
            _bstr_t iniFileName = _T("D:\\forum\\\\dwgExport.ini");

            VARIANT varProtType;
            varProtType.vt = VT_BSTR;
            varProtType.bstrVal = iniFileName;

            Result = options->put_Value(_T("Export_Acad_IniFile"), varProtType);
        }

        _bstr_t dwgFileName("D:\\forum\\Drawing1_AutoCAD.dwg");
        Result = dataMedium->put_FileName(dwgFileName);

        Result = translator->SaveCopyAs(doc, context, options, dataMedium);
    }

private:
    CComPtr<TranslatorAddIn> GetDwgTranslator() {
        HRESULT Result = NOERROR;

        _bstr_t ClientId = _T("{C24E3AC2-122E-11D5-8E91-0010B541CD80}");

        CComPtr<ApplicationAddIn> addin;
        Result = ThisApplication->ApplicationAddIns->get_ItemById(ClientId, &addin);

        CComPtr<TranslatorAddIn> translator;
        translator = addin;

        return translator;
    }
};
```

```vb
class ThisRule { 
public:
    CComPtr<Application> ThisApplication;

    void Main() {

        HRESULT Result = NOERROR;

        CComPtr<Document> gDoc;
        Result = ThisApplication->get_ActiveDocument(&gDoc);

        CComPtr<DrawingDocument> doc;
        doc = gDoc;
        
        CComPtr<TranslationContext> context;
        Result = ThisApplication->TransientObjects->CreateTranslationContext(&context);
        context->put_Type(kFileBrowseIOMechanism);

        CComPtr<NameValueMap> options;
        Result = ThisApplication->TransientObjects->CreateNameValueMap(&options);

        CComPtr<DataMedium> dataMedium;
        Result = ThisApplication->TransientObjects->CreateDataMedium(&dataMedium);

        CComPtr<TranslatorAddIn> translator;
        translator = GetDwgTranslator();


        if (translator->HasSaveCopyAsOptions[doc][context][options]) {
            _bstr_t iniFileName = _T("D:\\forum\\\\dwgExport.ini");

            VARIANT varProtType;
            varProtType.vt = VT_BSTR;
            varProtType.bstrVal = iniFileName;

            Result = options->put_Value(_T("Export_Acad_IniFile"), varProtType);
        }

        _bstr_t dwgFileName("D:\\forum\\Drawing1_AutoCAD.dwg");
        Result = dataMedium->put_FileName(dwgFileName);

        Result = translator->SaveCopyAs(doc, context, options, dataMedium);
    }

private:
    CComPtr<TranslatorAddIn> GetDwgTranslator() {
        HRESULT Result = NOERROR;

        _bstr_t ClientId = _T("{C24E3AC2-122E-11D5-8E91-0010B541CD80}");

        CComPtr<ApplicationAddIn> addin;
        Result = ThisApplication->ApplicationAddIns->get_ItemById(ClientId, &addin);

        CComPtr<TranslatorAddIn> translator;
        translator = addin;

        return translator;
    }
};
```

```vb
ThisRule rule;
rule.ThisApplication = GetInventor();
rule.Main();
```

```vb
ThisRule rule;
rule.ThisApplication = GetInventor();
rule.Main();
```

```vb
Sub Main()
    If Not ConfirmProceed() Then Return

    Dim sourceFolder As String = GetFolder("Select folder containing DWG files")
    If sourceFolder = "" Then Return

    Dim exportFolder As String = GetFolder("Select folder to save AutoCAD DWG files")
    If exportFolder = "" Then Return

    Dim configPath As String = "C:\Users\thoma\Documents\$_INDUSTRIAL_DRAFTING_&_DESIGN\INVENTOR_FILES\PRESETS\EXPORT-DWG.ini"
    If Not System.IO.File.Exists(configPath) Then
        MessageBox.Show("Missing DWG config: " & configPath)
        Return
    End If

    Dim files() As String = System.IO.Directory.GetFiles(sourceFolder, "*.dwg")
    Dim successCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim errorLog As New System.Text.StringBuilder
    Dim exportLog As New System.Text.StringBuilder

    exportLog.AppendLine("DWG Export Log - " & Now.ToString())
    exportLog.AppendLine("")

    For Each file As String In files
        Dim doc As Document = Nothing
        Try
            doc = ThisApplication.Documents.Open(file, False)
            If doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                doc.Close(True)
                Continue For
            End If

            Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(file)
            Dim drawDoc As DrawingDocument = doc
            Dim sheet1 As Sheet = drawDoc.Sheets(1)
            If sheet1.DrawingViews.Count = 0 Then Throw New Exception("No views on Sheet 1")
            Dim baseView As DrawingView = sheet1.DrawingViews(1)

            Dim modelDoc As Document = baseView.ReferencedDocumentDescriptor.ReferencedDocument
            If modelDoc Is Nothing Then Throw New Exception("Referenced model is missing.")

            Dim modelRev As String = ""
            Try
                modelRev = iProperties.Value(modelDoc, "Project", "Revision Number")
            Catch
                modelRev = "NO-REV"
            End Try

            Dim exportName As String = baseName & "-REV-" & SanitizeFileName(modelRev) & ".dwg"
            Dim exportPath As String = System.IO.Path.Combine(exportFolder, exportName)

            exportLog.AppendLine("DWG: " & baseName)
            exportLog.AppendLine(" ↳ Model: " & modelDoc.DisplayName)
            exportLog.AppendLine(" ↳ REV: " & modelRev)

            If System.IO.File.Exists(exportPath) Then
                exportLog.AppendLine(" ⏭ Skipped (same REV already exported)")
                doc.Close(True)
                skippedCount += 1
                Continue For
            End If

            ' Set up translator
            Dim addIn As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
            Dim context = ThisApplication.TransientObjects.CreateTranslationContext
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim options = ThisApplication.TransientObjects.CreateNameValueMap
            If addIn.HasSaveCopyAsOptions(doc, context, options) Then
                options.Value("Export_Acad_IniFile") = configPath
            End If
            Dim oData = ThisApplication.TransientObjects.CreateDataMedium
            oData.FileName = exportPath

            addIn.SaveCopyAs(doc, context, options, oData)
            doc.Close(True)
            successCount += 1
            exportLog.AppendLine(" ✅ Exported: " & exportName)
        Catch ex As Exception
            If Not doc Is Nothing Then doc.Close(True)
            errorLog.AppendLine("ERROR processing " & file & ": " & ex.Message)
        End Try

        exportLog.AppendLine("")
    Next

    ' Write logs
    System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "DWG_Export_Log.txt"), exportLog.ToString())
    If errorLog.Length > 0 Then
        System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "DWG_Export_Errors.txt"), errorLog.ToString())
    End If

    MessageBox.Show("DWG export complete." & vbCrLf & _
        "✔ Files exported: " & successCount & vbCrLf & _
        "⏭ Skipped (already exported at this REV): " & skippedCount & vbCrLf & _
        "⚠ Errors: " & errorLog.Length, "Export Summary")
End Sub

Function GetFolder(prompt As String) As String
    Dim dlg As New FolderBrowserDialog
    dlg.Description = prompt
    If dlg.ShowDialog() <> DialogResult.OK Then Return ""
    Return dlg.SelectedPath
End Function

Function ConfirmProceed() As Boolean
    Dim result = MessageBox.Show("Export DWGs only if target REV file does not exist?", "Confirm Export", MessageBoxButtons.YesNo)
    Return result = DialogResult.Yes
End Function

Function SanitizeFileName(input As String) As String
    For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
        input = input.Replace(c, "_")
    Next
    Return input
End Function
```

---

## Loading an Icon in Inventor 2025

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/loading-an-icon-in-inventor-2025/td-p/12745530](https://forums.autodesk.com/t5/inventor-programming-forum/loading-an-icon-in-inventor-2025/td-p/12745530)

**Author:** pfk

**Date:** ‎05-01-2024
	
		
		02:06 AM

**Description:** Hi My Inventor 2024 addin uses the following to put a button with an icon on the toolbar.              stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Supp...

**Code:**

```vb
stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(largeIcon);
#pragma warning restore CS0618 // Type or member is obsolete
            }

            mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                displayName, internalName, commandType,
                clientId, description, tooltip,
                standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);
```

```vb
stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(largeIcon);
#pragma warning restore CS0618 // Type or member is obsolete
            }

            mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                displayName, internalName, commandType,
                clientId, description, tooltip,
                standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);
```

---

## Run iLogic rule from Zero Document ribbon state

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/run-ilogic-rule-from-zero-document-ribbon-state/td-p/6840689#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/run-ilogic-rule-from-zero-document-ribbon-state/td-p/6840689#messageview_0)

**Author:** BrandonBG

**Date:** ‎01-30-2017
	
		
		08:15 AM

**Description:** Any suggestions on how to run an external iLogic rule from a ribbon button macro in the Zero Document state? I'm getting stuck with the VBA line: iLogicAuto.RunExternalRule oDoc, ExternalRuleNameIs it possible to run rules without referencing a document? Brandon

**Code:**

```vb
iLogicAuto.RunExternalRule oDoc, ExternalRuleName
```

```vb
ThisDoc.Launch("C:\TEMP\Part1.ipt")
oEditDoc = ThisApplication.ActiveEditDocument
MessageBox.Show(oEditDoc.FullFileName, "iLogic")
oEditDoc.Close
```

```vb
iLogicVb.RunExternalRule("ruleFileName")
```

---

## Add leading zeros to title block

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111#messageview_0)

**Author:** ludmylla.camielli

**Date:** ‎08-31-2020
	
		
		06:06 PM

**Description:** Hi guys, I need to fill my drawing's title block with the following, which will be the document number: XXXX.XXX.XXX.<Number of Sheets><Sheet Number> However, both numbers must have a leading zero (for example 0501, meaning 5 sheets and I'm currently looking at sheet 1). I have spent a couple of hours looking for any solution to this and ended making the following iLogic code, which I expected to be useful somehow. But I have no idea how to sync this with my document number. Is this the best way...

**Code:**

```vb
Dim oSheet As Sheet
Dim SheetNumber As Double

For Each oSheet In ThisApplication.ActiveDocument.Sheets
SheetNumber  = Mid(oSheet.Name, InStr(1, oSheet.Name, ":") + 1)
MessageBox.Show(Microsoft.VisualBasic.Strings.Format(SheetNumber,"0#"), oSheet.Name)
Next
```

```vb
'Get Total Number of Sheets
SheetCount = ThisApplication.ActiveDocument.Sheets.Count

'Append a leading zero if less than 10
If SheetCount < 10 Then
	SheetCo = "0" & SheetCount
Else 
	SheetCo = SheetCount
End If

'Loop through each sheet and grab the sheet number from the name
For Each Sheet In ThisApplication.ActiveDocument.Sheets
SheetNumber = Mid(Sheet.Name, InStr(1, Sheet.Name, ":") + 1)

'Append a leading zero if less than 10
If SheetNumber <10 Then 
SheetNo = "0" & SheetNumber
Else
	SheetNo = SheetNumber
End If

MessageBox.Show(SheetCo & SheetNo, "Sheet descriptor")

Next
```

```vb
'Get Total Number of Sheets
SheetCount = ThisApplication.ActiveDocument.Sheets.Count

'Append a leading zero if less than 10
If SheetCount < 10 Then
	SheetCo = "0" & SheetCount
Else 
	SheetCo = SheetCount
End If

'Loop through each sheet and grab the sheet number from the name
For Each Sheet In ThisApplication.ActiveDocument.Sheets
SheetNumber = Mid(Sheet.Name, InStr(1, Sheet.Name, ":") + 1)

'Append a leading zero if less than 10
If SheetNumber <10 Then 
SheetNo = "0" & SheetNumber
Else
	SheetNo = SheetNumber
End If

MessageBox.Show(SheetCo & SheetNo, "Sheet descriptor")

Next
```

```vb
Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)).ToString("00")
	MsgBox(oSheets.Count.ToString("00") & oSheetNumber)
Next
```

```vb
Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)).ToString("00")
	Dim oTB As TitleBlock = oSheet.TitleBlock
	For Each oTextBox As Inventor.TextBox In oTB.Definition.Sketch.TextBoxes
		If oTextBox.Text = "<CustomSheetNumber>"
			oTB.SetPromptResultText(oTextBox, (oSheets.Count.ToString("00") & oSheetNumber))
			Exit For
		End If
	Next
Next
```

```vb
Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)+SheetStartNum-1).ToString("0000")
	
	Dim oTB As TitleBlock = oSheet.TitleBlock
	For Each oTextBox As Inventor.TextBox In oTB.Definition.Sketch.TextBoxes
		If oTextBox.Text = "<CustomSheetNumber>"
			oTB.SetPromptResultText(oTextBox, (oSheetNumber))
			Exit For
		End If
	Next
Next
```

---

## Rule with multiple options

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705](https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705)

**Author:** KaynePellegrino

**Date:** ‎08-26-2025
	
		
		09:28 AM

**Description:** This is a long one.------------------------ I make doors. They can be different Widths and Heights, among different Types.Widths:3-6 (41.75")3-0 (35.75")2-10 (33.75")2-8 (31.75")2-6 (29.75") Heights:8-0 (95.375")6-8 (79.375") Types:1K (1-Panel)2K (2-Panel3F (3-Panel)3S (3-Panel w/ Glass)50 (5-Panel) ------------------------ I have a Width Param & Rule:DoorWidthOptions [Text, Multi]: 3-6, 3-0, 2-10, 2-8, 2-6DoorWidth [Num]: (Varies)Door Width Rule:'Door Width
If DoorWidthOptions = "3-6" Then
		Do...

**Code:**

```vb
'Door Width
If DoorWidthOptions = "3-6" Then
		DoorWidth = 41.75
	End If

If DoorWidthOptions = "3-0" Then
		DoorWidth = 35.75
	End If

If DoorWidthOptions = "2-10" Then
		DoorWidth = 33.75
	End If

If DoorWidthOptions = "2-8" Then
		DoorWidth = 31.75
	End If

If DoorWidthOptions = "2-6" Then
		DoorWidth = 29.75
	End If
```

```vb
'Door Height
If DoorHeightOptions = "6-8" Then
		DoorHeight = 79.375
	End If

If DoorHeightOptions = "8-0" Then
		DoorHeight = 95.375
	End If
```

```vb
If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
		Component.IsActive("2K Skin: Back") = False
		Component.IsActive("2K Skin: Front") = False
		Component.IsActive("3F Skin: Back") = False
		Component.IsActive("3F Skin: Front") = False
		Component.IsActive("3F (3-6) Skin: Back") = False
		Component.IsActive("3F (3-6) Skin: Front") = False
		Component.IsActive("3S Skin: Back") = False
		Component.IsActive("3S Skin: Front") = False
		Component.IsActive("50 Skin: Back") = False
		Component.IsActive("50 Skin: Front") = False
	End If
```

```vb
If DoorType = "3F" And DoorWidthOptions = ("2-6" or "2-8" or "2-10" or "3-0") Then
		Component.IsActive("1K Skin: Back") = False
		Component.IsActive("1K Skin: Front") = False
		Component.IsActive("2K Skin: Back") = False
		Component.IsActive("2K Skin: Front") = False
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
		Component.IsActive("3F (3-6) Skin: Back") = False
		Component.IsActive("3F (3-6) Skin: Front") = False
		Component.IsActive("3S Skin: Back") = False
		Component.IsActive("3S Skin: Front") = False
		Component.IsActive("50 Skin: Back") = False
		Component.IsActive("50 Skin: Front") = False
	End If
```

```vb
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False
If DoorType = "3F" And DoorWidthO >= 29.75 And DoorWidthO <= 35.75
	Component.IsActive("3F Skin: Back") = True
	Component.IsActive("3F Skin: Front") = True
ElseIf DoorType = "3S" And DoorWidthO >= 29.75 And DoorWidthO <= 35.75
	Component.IsActive("3S Skin: Back") = True
	Component.IsActive("3S Skin: Front") = True
End If
```

```vb
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False
Select Case DoorType
Case "3F"
	Select Case DoorWidthO
	Case 29.75 To 35.75
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True	
	Case 2 To 29.75
		'Etc Etc Etc
	Case Else
		'Etc Etc Etc
	End Select
Case "3S"
	Select Case DoorWidthO
	Case 12 To 20
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case 2 To 12
		'Etc Etc Etc
	End Select
End Select
```

```vb
'"3F Skin" ActiveIf DoorType = "3F" And DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If
'"3F (3-6) Skin" Active
If DoorType = "3F" And DoorWidthOptions = "3-6" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
```

```vb
Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F"
	Select Case DoorWidthOptions
	Case 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 22
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 20
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	Case 18
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select
```

```vb
Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F"
	Select Case DoorWidthOptions
	Case 18 to 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select
```

```vb
Dim DoorWidthOptions As Double = '???
Select Case DoorType
Case "3F", "3F Skin", "3F (3-6) Skin"
	Select Case DoorWidthOptions
	Case 18, 20, 22, 36
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End Select
End Select
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	End If

If DoorType = "2K" Then
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If

If DoorType = "3S" Then
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	End If

If DoorType = "50" Then
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	End If
```

```vb
If DoorType = "3F" And DoorWidthOptions = "3-0" or "2-10" or "2-8" or "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End IfIf DoorType = "3F" And DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then                Component.IsActive("3F Skin: Back") = True                Component.IsActive("3F Skin: Front") = True        End If
```

```vb
'Door Height
If DoorHeightOptions = "6-8" Then
		DoorHeight = 79.375
	End If

If  DoorHeightOptions = "8-0" Then
		DoorHeight = 95.375
	End If


'Door Width
If DoorWidthOptions = "3-6" Then
		DoorWidth = 41.75
	End If

If DoorWidthOptions = "3-0" Then
		DoorWidth = 35.75
	End If

If DoorWidthOptions = "2-10" Then
		DoorWidth = 33.75
	End If

If DoorWidthOptions = "2-8" Then
		DoorWidth = 31.75
	End If

If DoorWidthOptions = "2-6" Then
		DoorWidth = 29.75
	End If
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

If DoorType = "1K" Then
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	End If

If DoorType = "2K" Then
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth <= 40 Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth >= 40  And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If

If DoorType = "3S" Then
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	End If

If DoorType = "50" Then
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	End If
```

```vb
If DoorType = "3F" And DoorWidth <= 40 Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	End If

If DoorType = "3F" And DoorWidth >= 40  And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
```

```vb
If DoorType = "3F" Then
	If DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
ElseIf DoorType = "X" Then
	'Etc etc etc
ElseIf DoorType = "Y" Then
	'Etc etc etc
ElseIf DoorType = "Z" Then
	'Etc etc etc
End If
```

```vb
If DoorType = "3F" And DoorWidth = 35.75, 33.75, 31.75, AndOr 29.75 Then                "Component1" = True        ElseIf DoorWidth = 41.75 Then                "Component2" = True        End If
```

```vb
If "Variable1" = "OptionA1" And "Variable2" = "OptionB1" OR "OptionB2" Then '(Multiple Options in the Same Line, separated by some kind of 'or' statement)
		"Component1" = True                Else If "Variable2" = "OptionB3" Then                "Component2" = True
	End If
```

```vb
If DoorType = "3F" Then
	If DoorWidthOptions = "3-0" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-10" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-8" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = "2-6" Then
		Component.IsActive("3F Skin: Back") = True
		Component.IsActive("3F Skin: Front") = True
	ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
		Component.IsActive("3F (3-6) Skin: Back") = True
		Component.IsActive("3F (3-6) Skin: Front") = True
	End If
```

```vb
Sub main()
Dim X as integer = 1
Dim X1 as integer = 2 

Dim Y as integer = 3
Dim Y1 as integer = 4

If (x = x1 and y = y1) or y = 4 then
     'do something
Else if X = Y then 
     'Do something else
Else

End if 





End sub
```

```vb
Sub main()
Dim X as integer = 1
Dim X1 as integer = 2 

Dim Y as integer = 3
Dim Y1 as integer = 4

If (x = x1 and y = y1) or y = 4 then
     'do something
Else if X = Y then 
     'Do something else
Else

End if 





End sub
```

```vb
If (MiddleLetter.StartsWith("D") = True Or MiddleLetter.StartsWith("K") = True) And LastLetter.StartsWith("H") = False Then 
				Tank_Leg_Stand_TC = True	
			Else
				Tank_Leg_Stand_TC = False
			End If
```

```vb
If (MiddleLetter.StartsWith("D") = True Or MiddleLetter.StartsWith("K") = True) And LastLetter.StartsWith("H") = False Then 
				Tank_Leg_Stand_TC = True	
			Else
				Tank_Leg_Stand_TC = False
			End If
```

```vb
Dim A,B,C As String
A = "1"
B = "0"
C = "0"

If A = "1" And (B = "1" Or C = "1")
	MessageBox.Show("1", "Test")
Else
	MessageBox.Show("2", "Test")
end if
```

```vb
ElseIf DoorWidthOptions = DoorWidthOptions = "3-6" And DoorHeightOptions = "8-0" Then
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

Select Case DoorType
	Case "1K" 
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	Case "2K" 
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	Case "3F"
		If DoorWidthOptions <> "3-6"
			Component.IsActive("3F Skin: Back") = True
			Component.IsActive("3F Skin: Front") = True
		Else 
			If DoorHeightOptions = "8-0"
				Component.IsActive("3F (3-6) Skin: Back") = True
				Component.IsActive("3F (3-6) Skin: Front") = True
			Else
				MessageBox.Show(String.Format("DoorType = {0} | DoorWidthOptions = {1} | DoorHeightOptions = {2}", DoorType, DoorWidthOptions, DoorHeightOptions).ToString, "Un-Accounted For Result:")
			End If
		End If
	Case "3S" 
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case "50" 
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	Case Else
		MessageBox.Show(String.Format("DoorType = {0}", DoorType).ToString, "Un-Accounted For DoorType:")
End Select
```

```vb
'Door Type
Component.IsActive("1K Skin: Back") = False
Component.IsActive("1K Skin: Front") = False
Component.IsActive("2K Skin: Back") = False
Component.IsActive("2K Skin: Front") = False
Component.IsActive("3F Skin: Back") = False
Component.IsActive("3F Skin: Front") = False
Component.IsActive("3F (3-6) Skin: Back") = False
Component.IsActive("3F (3-6) Skin: Front") = False
Component.IsActive("3S Skin: Back") = False
Component.IsActive("3S Skin: Front") = False
Component.IsActive("50 Skin: Back") = False
Component.IsActive("50 Skin: Front") = False

Select Case DoorType
	Case "1K" 
		Component.IsActive("1K Skin: Back") = True
		Component.IsActive("1K Skin: Front") = True
	Case "2K" 
		Component.IsActive("2K Skin: Back") = True
		Component.IsActive("2K Skin: Front") = True
	Case "3F"
		If DoorWidthOptions <> "3-6"
			Component.IsActive("3F Skin: Back") = True
			Component.IsActive("3F Skin: Front") = True
		Else 
			If DoorHeightOptions = "8-0"
				Component.IsActive("3F (3-6) Skin: Back") = True
				Component.IsActive("3F (3-6) Skin: Front") = True
			Else
				MessageBox.Show(String.Format("DoorType = {0} | DoorWidthOptions = {1} | DoorHeightOptions = {2}", DoorType, DoorWidthOptions, DoorHeightOptions).ToString, "Un-Accounted For Result:")
			End If
		End If
	Case "3S" 
		Component.IsActive("3S Skin: Back") = True
		Component.IsActive("3S Skin: Front") = True
	Case "50" 
		Component.IsActive("50 Skin: Back") = True
		Component.IsActive("50 Skin: Front") = True
	Case Else
		MessageBox.Show(String.Format("DoorType = {0}", DoorType).ToString, "Un-Accounted For DoorType:")
End Select
```

```vb
If DoorWidthOptions <> "3-6"
```

```vb
Active_Condition = (DoorType = "3F" And (DoorWidthOptions = "2-6" Or DoorWidthOptions = "2-8" Or DoorWidthOptions = "2-10" Or DoorWidthOptions = "3-0"))

Component.IsActive("1K Skin: Back") = Active_Condition
Component.IsActive("1K Skin: Front") = Active_Condition
Component.IsActive("2K Skin: Back") = Not Active_Condition
Component.IsActive("2K Skin: Front") = Not Active_Condition
Component.IsActive("3F Skin: Back") = Not Active_Condition
Component.IsActive("3F Skin: Front") = Not Active_Condition
Component.IsActive("3F (3-6) Skin: Back") = Not Active_Condition
Component.IsActive("3F (3-6) Skin: Front") = Not Active_Condition
Component.IsActive("3S Skin: Back") = Not Active_Condition
Component.IsActive("3S Skin: Front") = Not Active_Condition
Component.IsActive("50 Skin: Back") = Not Active_Condition
Component.IsActive("50 Skin: Front") = Not Active_Condition
```

---

## Using Ilogic to update model states

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/using-ilogic-to-update-model-states/td-p/12844817#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/using-ilogic-to-update-model-states/td-p/12844817#messageview_0)

**Author:** aberningR89Y3

**Date:** ‎06-17-2024
	
		
		10:52 AM

**Description:** I am looking for a way to update every model state in a part file without having to manually select each model state and click 'local update'.Ideally, I would like to do this without having to know the name of every model state beforehand. Something like: find each model state name, put it into an array, and then loop through and useThisApplication.ActiveView.Update()I am not well versed on VBA/iLogic syntax so maybe there is a 'For Each' loop that can do it?   Thanks in advance,Alex
					
				
...

**Code:**

```vb
ThisApplication.ActiveView.Update()
```

```vb
Dim oDoc As Inventor.Document = ThisDoc.FactoryDocument
If oDoc Is Nothing Then Return
Dim oMSs As ModelStates = oDoc.ComponentDefinition.ModelStates
If oMSs Is Nothing OrElse oMSs.Count = 0 Then Return
Dim oAMS As ModelState = oMSs.ActiveModelState
For Each oMS As ModelState In oMSs
	oMS.Activate
	Dim oMSDoc As Inventor.Document = TryCast(oMS.Document, Inventor.Document)
	If oMSDoc Is Nothing Then Continue For
	oMSDoc.Update2(True)
Next oMS
If oMSs.ActiveModelState IsNot oAMS Then oAMS.Activate
```

```vb
Dim oDoc As Inventor.Document = ThisDoc.FactoryDocument
If oDoc Is Nothing Then Return
Dim oMSs As ModelStates = oDoc.ComponentDefinition.ModelStates
If oMSs Is Nothing OrElse oMSs.Count = 0 Then Return
Dim oAMS As ModelState = oMSs.ActiveModelState
For Each oMS As ModelState In oMSs
	oMS.Activate
	Dim oMSDoc As Inventor.Document = TryCast(oMS.Document, Inventor.Document)
	If oMSDoc Is Nothing Then Continue For
	oMSDoc.Update2(True)
Next oMS
If oMSs.ActiveModelState IsNot oAMS Then oAMS.Activate
```

---

## iLogic Code to export Sheets to IDWS

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-export-sheets-to-idws/td-p/13705767](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-export-sheets-to-idws/td-p/13705767)

**Author:** rossPJ9W7

**Date:** ‎07-01-2025
	
		
		05:12 AM

**Description:** I need help with this code. The goal is to export all the sheets in a idw file to seperate idws with the sheet name as the file name. The code works. All the sheets are saved in a new folder created called " Sheets". All new idw files are named with the sheet name. Issue: The new idws dont open. From my understanding the way the code needs to work is, add a number on the sheet names to be a counter. Then deleted all sheets except one and save as new idw file. Then go to the next counter number a...

**Code:**

```vb
Sub Main
	'get the Inventor Application to a variable
	Dim oInvApp As Inventor.Application = ThisApplication
	'get the main Documents object of the application, for use later
	Dim oDocs As Inventor.Documents = oInvApp.Documents
	'get the 'active' Document, and try to 'Cast' it to the DrawingDocument Type
	'if this fails, then no value will be assigned to the variable
	'but no error will happen either
	Dim oDDoc As DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	'check if the variable got assigned a value...if not, then exit rule
	If oDDoc Is Nothing Then Return
	'get the main Sheets collection of the 'active' drawing
	Dim oSheets As Inventor.Sheets = oDDoc.Sheets
	'if only 1 sheet, then exit rule, because noting to do
	If oSheets.Count = 1 Then Return
	'record currently 'active' Sheet, to reset to active at the end
	Dim oASheet As Inventor.Sheet = oDDoc.ActiveSheet
	'get the full path of the active drawing, without final directory separator character
	Dim sPath As String = System.IO.Path.GetDirectoryName(oDDoc.FullFileName)
	'specify name of sub folder to put new drawing documents into for each sheet
	Dim sSubFolder As String = "Sheets"
	'combine main Path with sub folder name, to make new path
	'adds the directory separator character in for us, automatically
	Dim sNewPath As String = System.IO.Path.Combine(sPath, sSubFolder)
	'create the sub directory, if it does not already exist
	If Not System.IO.Directory.Exists(sNewPath) Then
		System.IO.Directory.CreateDirectory(sNewPath)
	End If
	'not using 'For Each' on purpose, to maintain proper order 
	For iSheetIndex As Integer = 1 To oSheets.Count
		'get the Sheet at this index
		Dim oSheet As Inventor.Sheet = oSheets.Item(iSheetIndex)
		'make this sheet the active sheet
		oSheet.Activate()
		Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
		'save copy of this drawing to new file, in background (new copy not open yet)
		oDDoc.SaveAs(sNewDrawingFFN, True)
		'open that new drawing file (invisibly), so we can modify it
		Dim oNewDDoc As DrawingDocument = oDocs.Open(sNewDrawingFFN, False)
		'get Sheets collection of new drawing to a variable
		Dim oNewDocSheets As Inventor.Sheets = oNewDDoc.Sheets
		'get the Sheet at same Index in new drawing (the one to keep)
		Dim oThisNewDocSheet As Inventor.Sheet = oNewDocSheets.Item(iSheetIndex)
		'delete all other sheets in new drawing besides this sheet
		For Each oNewDocSheet As Inventor.Sheet In oNewDocSheets
			If Not oNewDocSheet Is oThisNewDocSheet Then
				Try
					'Optional 'RetainDependentViews' is False by default
					oNewDocSheet.Delete() 
				Catch
					Logger.Error("Error deleting Sheet named:  " & oNewDocSheet.Name)
				End Try
			End If
		Next
		'save new drawing document again, after changes
		oNewDDoc.Save()
		'next line only valid when not visibly opened, and not being referenced
		oNewDDoc.ReleaseReference()
		'if visibly opened, then use oNewDDoc.Close() method instead
	Next 'oSheet
	'activate originally active sheet again
	oASheet.Activate()
	'only clear out all Documents that are 'unreferenced' in Inventor's memory
	'used after ReleaseReference, to clean-up
	oDocs.CloseAll(True)
End Sub
```

```vb
Sub Main
	'get the Inventor Application to a variable
	Dim oInvApp As Inventor.Application = ThisApplication
	'get the main Documents object of the application, for use later
	Dim oDocs As Inventor.Documents = oInvApp.Documents
	'get the 'active' Document, and try to 'Cast' it to the DrawingDocument Type
	'if this fails, then no value will be assigned to the variable
	'but no error will happen either
	Dim oDDoc As DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	'check if the variable got assigned a value...if not, then exit rule
	If oDDoc Is Nothing Then Return
	'get the main Sheets collection of the 'active' drawing
	Dim oSheets As Inventor.Sheets = oDDoc.Sheets
	'if only 1 sheet, then exit rule, because noting to do
	If oSheets.Count = 1 Then Return
	'record currently 'active' Sheet, to reset to active at the end
	Dim oASheet As Inventor.Sheet = oDDoc.ActiveSheet
	'get the full path of the active drawing, without final directory separator character
	Dim sPath As String = System.IO.Path.GetDirectoryName(oDDoc.FullFileName)
	'specify name of sub folder to put new drawing documents into for each sheet
	Dim sSubFolder As String = "Sheets"
	'combine main Path with sub folder name, to make new path
	'adds the directory separator character in for us, automatically
	Dim sNewPath As String = System.IO.Path.Combine(sPath, sSubFolder)
	'create the sub directory, if it does not already exist
	If Not System.IO.Directory.Exists(sNewPath) Then
		System.IO.Directory.CreateDirectory(sNewPath)
	End If
	'not using 'For Each' on purpose, to maintain proper order 
	For iSheetIndex As Integer = 1 To oSheets.Count
		'get the Sheet at this index
		Dim oSheet As Inventor.Sheet = oSheets.Item(iSheetIndex)
		'make this sheet the active sheet
		oSheet.Activate()
		Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
		'save copy of this drawing to new file, in background (new copy not open yet)
		oDDoc.SaveAs(sNewDrawingFFN, True)
		'open that new drawing file (invisibly), so we can modify it
		Dim oNewDDoc As DrawingDocument = oDocs.Open(sNewDrawingFFN, False)
		'get Sheets collection of new drawing to a variable
		Dim oNewDocSheets As Inventor.Sheets = oNewDDoc.Sheets
		'get the Sheet at same Index in new drawing (the one to keep)
		Dim oThisNewDocSheet As Inventor.Sheet = oNewDocSheets.Item(iSheetIndex)
		'delete all other sheets in new drawing besides this sheet
		For Each oNewDocSheet As Inventor.Sheet In oNewDocSheets
			If Not oNewDocSheet Is oThisNewDocSheet Then
				Try
					'Optional 'RetainDependentViews' is False by default
					oNewDocSheet.Delete() 
				Catch
					Logger.Error("Error deleting Sheet named:  " & oNewDocSheet.Name)
				End Try
			End If
		Next
		'save new drawing document again, after changes
		oNewDDoc.Save()
		'next line only valid when not visibly opened, and not being referenced
		oNewDDoc.ReleaseReference()
		'if visibly opened, then use oNewDDoc.Close() method instead
	Next 'oSheet
	'activate originally active sheet again
	oASheet.Activate()
	'only clear out all Documents that are 'unreferenced' in Inventor's memory
	'used after ReleaseReference, to clean-up
	oDocs.CloseAll(True)
End Sub
```

```vb
iSheetIndex.ToString()
```

```vb
Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
```

```vb
Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
```

```vb
'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
```

```vb
'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
```

```vb
Sub Main
	'get the Inventor Application to a variable
	Dim oInvApp As Inventor.Application = ThisApplication
	'get the main Documents object of the application, for use later
	Dim oDocs As Inventor.Documents = oInvApp.Documents
	'get the 'active' Document, and try to 'Cast' it to the DrawingDocument Type
	'if this fails, then no value will be assigned to the variable
	'but no error will happen either
	Dim oDDoc As DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	'check if the variable got assigned a value...if not, then exit rule
	If oDDoc Is Nothing Then Return
	'get the main Sheets collection of the 'active' drawing
	Dim oSheets As Inventor.Sheets = oDDoc.Sheets
	'if only 1 sheet, then exit rule, because noting to do
	If oSheets.Count = 1 Then Return
	'record currently 'active' Sheet, to reset to active at the end
	Dim oASheet As Inventor.Sheet = oDDoc.ActiveSheet
	'get the full path of the active drawing, without final directory separator character
	Dim sPath As String = System.IO.Path.GetDirectoryName(oDDoc.FullFileName)
	'specify name of sub folder to put new drawing documents into for each sheet
	Dim sSubFolder As String = "Sheets"
	'combine main Path with sub folder name, to make new path
	'adds the directory separator character in for us, automatically
	Dim sNewPath As String = System.IO.Path.Combine(sPath, sSubFolder)
	'create the sub directory, if it does not already exist
	If Not System.IO.Directory.Exists(sNewPath) Then
		System.IO.Directory.CreateDirectory(sNewPath)
	End If
	'not using 'For Each' on purpose, to maintain proper order 
	For iSheetIndex As Integer = 1 To oSheets.Count
		'get the Sheet at this index
		Dim oSheet As Inventor.Sheet = oSheets.Item(iSheetIndex)
		'make this sheet the active sheet
		oSheet.Activate()
		Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		'sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
		'save copy of this drawing to new file, in background (new copy not open yet)
		oDDoc.SaveAs(sNewDrawingFFN , True)
		'open that new drawing file (invisibly), so we can modify it
		Dim oNewDDoc As DrawingDocument = oDocs.Open(sNewDrawingFFN, False)
		'get Sheets collection of new drawing to a variable
		Dim oNewDocSheets As Inventor.Sheets = oNewDDoc.Sheets
		'get the Sheet at same Index in new drawing (the one to keep)
		Dim oThisNewDocSheet As Inventor.Sheet = oNewDocSheets.Item(iSheetIndex)
		'delete all other sheets in new drawing besides this sheet
		For Each oNewDocSheet As Inventor.Sheet In oNewDocSheets
			If Not oNewDocSheet Is oThisNewDocSheet Then
				Try
					'Optional 'RetainDependentViews' is False by default
					oNewDocSheet.Delete() 
				Catch
					Logger.Error("Error deleting Sheet named:  " & oNewDocSheet.Name)
				End Try
			End If
		Next
		'save new drawing document again, after changes
		oNewDDoc.Save()
		'next line only valid when not visibly opened, and not being referenced
		oNewDDoc.ReleaseReference()
		'if visibly opened, then use oNewDDoc.Close() method instead
	Next 'oSheet
	'activate originally active sheet again
	oASheet.Activate()
	'only clear out all Documents that are 'unreferenced' in Inventor's memory
	'used after ReleaseReference, to clean-up
	oDocs.CloseAll(True)
End Sub
```

---

## VBA : Using GetObject in Inventor 2024

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/vba-using-getobject-in-inventor-2024/td-p/13777048#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/vba-using-getobject-in-inventor-2024/td-p/13777048#messageview_0)

**Author:** TONELLAL

**Date:** ‎08-21-2025
	
		
		04:05 AM

**Description:** Hello,I have several macros in VBA, using Excel. They connect to Excel using GetObject(, "Excel.application"), or CreateObject("Excel.application").These commands are not existing anymore in .NET framework > 5, so they don't work anymore with Inventor 2024.I found several links using .NET or C#, but noting on VBA.How can I modify these macros, still in VBA, so that they work with version 2024 and future versions of Inventor?

**Code:**

```vb
Option Explicit
Sub OpenExcel()
    Dim oExcelApp As Object
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible = True
    
    Dim oWB As Object
    Set oWB = oExcelApp.Workbooks.Add()
    
    Dim oWS As Object
    Set oWS = oWB.Worksheets.Add()
    oWS.Name = "My New Worksheet"
    
    Call MsgBox("Review New Instance Of Excel, And New, Renamed Sheet In It.", , "")
    
    Call oWB.Close
    Call oExcelApp.Quit
    
    Set oWS = Nothing
    Set oWB = Nothing
    Set oExcelApp = Nothing
End Sub
```

```vb
Option Explicit
Sub OpenExcel()
    Dim oExcelApp As Object
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible = True
    
    Dim oWB As Object
    Set oWB = oExcelApp.Workbooks.Add()
    
    Dim oWS As Object
    Set oWS = oWB.Worksheets.Add()
    oWS.Name = "My New Worksheet"
    
    Call MsgBox("Review New Instance Of Excel, And New, Renamed Sheet In It.", , "")
    
    Call oWB.Close
    Call oExcelApp.Quit
    
    Set oWS = Nothing
    Set oWB = Nothing
    Set oExcelApp = Nothing
End Sub
```

---

## Loading an Icon in Inventor 2025

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/loading-an-icon-in-inventor-2025/td-p/12745530#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/loading-an-icon-in-inventor-2025/td-p/12745530#messageview_0)

**Author:** pfk

**Date:** ‎05-01-2024
	
		
		02:06 AM

**Description:** Hi My Inventor 2024 addin uses the following to put a button with an icon on the toolbar.              stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Supp...

**Code:**

```vb
stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(largeIcon);
#pragma warning restore CS0618 // Type or member is obsolete
            }

            mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                displayName, internalName, commandType,
                clientId, description, tooltip,
                standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);
```

```vb
stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(largeIcon);
#pragma warning restore CS0618 // Type or member is obsolete
            }

            mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                displayName, internalName, commandType,
                clientId, description, tooltip,
                standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);
```

---

## Macro to delete specific sketch symbol?

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606](https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606)

**Author:** jacob_r_kane

**Date:** ‎08-26-2025
	
		
		08:37 AM

**Description:** Hello,Without spending days trying to learn this myself, what is the code to scan all drawing sheets, and delete a sketch symbol named rev triangle?Thanks
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, Inventor.DrawingDocument)
If oDDoc Is Nothing Then Return
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oSS As SketchedSymbol In oSheet.SketchedSymbols
		If oSS.Name = "rev triangle" Then
			oSS.Delete()
		End If
	Next 'oSS
Next 'oSheet
```

```vb
Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, Inventor.DrawingDocument)
If oDDoc Is Nothing Then Return
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oSS As SketchedSymbol In oSheet.SketchedSymbols
		If oSS.Name = "rev triangle" Then
			oSS.Delete()
		End If
	Next 'oSS
Next 'oSheet
```

```vb
Sub DeleteRevTriangleSketchedSymbols()
    If Not ThisApplication.ActiveDocumentType = kDrawingDocumentObject Then Return
    Dim oDDoc As DrawingDocument
    Set oDDoc = ThisApplication.ActiveDocument
    Dim oSheet As Inventor.Sheet
    For Each oSheet In oDDoc.Sheets
        Dim oSS As SketchedSymbol
        For Each oSS In oSheet.SketchedSymbols
            If oSS.Name = "rev triangle" Then
                Call oSS.Delete
            End If
        Next 'oSS
    Next 'oSheet
End Sub
```

```vb
Sub DeleteRevTriangleSketchedSymbols()
    If Not ThisApplication.ActiveDocumentType = kDrawingDocumentObject Then Return
    Dim oDDoc As DrawingDocument
    Set oDDoc = ThisApplication.ActiveDocument
    Dim oSheet As Inventor.Sheet
    For Each oSheet In oDDoc.Sheets
        Dim oSS As SketchedSymbol
        For Each oSS In oSheet.SketchedSymbols
            If oSS.Name = "rev triangle" Then
                Call oSS.Delete
            End If
        Next 'oSS
    Next 'oSheet
End Sub
```

```vb
Public Sub DeleteSpecificSketchSymbol()
    ' Set a reference to the active drawing document.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    ' Define the name of the sketch symbol to delete.
    Const sSymbolName As String = "Rev Triangle" ' Replace with the actual symbol name
    ' Loop through each sheet in the drawing document.
    Dim oSheet As Sheet
    For Each oSheet In oDrawDoc.Sheets
        ' Activate the current sheet (optional, but good practice for visibility).
        oSheet.Activate
        ' Loop through each sketched symbol on the current sheet.
        Dim oSketchedSymbol As SketchedSymbol
        Dim i As Long
        For i = oSheet.SketchedSymbols.Count To 1 Step -1 ' Loop backwards to avoid issues when deleting
            Set oSketchedSymbol = oSheet.SketchedSymbols.Item(i)
            ' Check if the symbol's name matches the target name.
            If oSketchedSymbol.Name = sSymbolName Then
                ' Delete the instance of the sketch symbol.
                oSketchedSymbol.Delete
            End If
        Next i
    Next oSheet
End Sub
```

```vb
Public Sub DeleteSpecificSketchSymbol()
    ' Set a reference to the active drawing document.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    ' Define the name of the sketch symbol to delete.
    Const sSymbolName As String = "Rev Triangle" ' Replace with the actual symbol name
    ' Loop through each sheet in the drawing document.
    Dim oSheet As Sheet
    For Each oSheet In oDrawDoc.Sheets
        ' Activate the current sheet (optional, but good practice for visibility).
        oSheet.Activate
        ' Loop through each sketched symbol on the current sheet.
        Dim oSketchedSymbol As SketchedSymbol
        Dim i As Long
        For i = oSheet.SketchedSymbols.Count To 1 Step -1 ' Loop backwards to avoid issues when deleting
            Set oSketchedSymbol = oSheet.SketchedSymbols.Item(i)
            ' Check if the symbol's name matches the target name.
            If oSketchedSymbol.Name = sSymbolName Then
                ' Delete the instance of the sketch symbol.
                oSketchedSymbol.Delete
            End If
        Next i
    Next oSheet
End Sub
```

---

## Excluding Print and Count if Sheet name matches

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697#messageview_0)

**Author:** adeel.malik

**Date:** ‎09-01-2022
	
		
		08:03 PM

**Description:** I have a bunch of sheets named Vendor (e.g. Vendor:1 or Vendor:2...). For certain things I want to exclude these sheets from count and print. I wanted to write an ilogic for this but my code is not working.  Dim oDoc As DrawingDocument
oDoc = ThisApplication.ActiveDocument
Dim oSheet As Sheet

i = 0
x = 0

For Each oSheet In oDoc.Sheets
	i = i + 1
Next

While (x<i)
	Sheet_Name = "Vendor:" & x 
	x= x+1
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromPrinting = True  <---Not working
	'ThisDo...

**Code:**

```vb
Dim oDoc As DrawingDocument
oDoc = ThisApplication.ActiveDocument
Dim oSheet As Sheet

i = 0
x = 0

For Each oSheet In oDoc.Sheets
	i = i + 1
Next

While (x<i)
	Sheet_Name = "Vendor:" & x 
	x= x+1
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromPrinting = True  <---Not working
	'ThisDoc.Document.Sheets.Item(Sheet_Name).ExcludeFromCount = True <---Not working
End While
```

```vb
Dim oDoc As DrawingDocument = ThisDoc.Document
Dim i as Integer = 0

For Each oSheet As Sheet In oDoc.Sheets
	
	i = i + 1
	
	If oSheet.Name = "Vendor:" & i
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
	End If
Next
```

```vb
Dim oDoc As DrawingDocument = ThisDoc.Document
Dim i as Integer = 0

For Each oSheet As Sheet In oDoc.Sheets
	
	i = i + 1
	
	If oSheet.Name = "Vendor:" & i
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
		oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
	End If
Next
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeVendorSheets As Boolean = InputRadioBox("Exclude Vendor Sheets?", "True", "False", False, "Toggle Vendor Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("Vendor") Then
		oSheet.ExcludeFromCount = oExcludeVendorSheets
		oSheet.ExcludeFromPrinting = oExcludeVendorSheets
	End If
Next
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeVendorSheets As Boolean = InputRadioBox("Exclude Vendor Sheets?", "True", "False", False, "Toggle Vendor Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("Vendor") Then
		oSheet.ExcludeFromCount = oExcludeVendorSheets
		oSheet.ExcludeFromPrinting = oExcludeVendorSheets
	End If
Next
```

```vb
If oSheet.Name.Contains("Vendor") Then
   oDoc.Sheets.Item(oSheet.Name).ExcludeFromPrinting = True  
   oDoc.Sheets.Item(oSheet.Name).ExcludeFromCount = True
End If
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.StartsWith("S") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next
```

```vb
If oSheet.Name.StartsWith("S") Then
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If oSheet.Name.NotStartsWith("C") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next
```

```vb
Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If Not oSheet.Name.StartsWith("C") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next
```

---

## Check Vault Status

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/check-vault-status/td-p/13834312#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/check-vault-status/td-p/13834312#messageview_0)

**Author:** karthur1

**Date:** ‎10-02-2025
	
		
		05:27 AM

**Description:** I have an iLogic rule to quickly print to PDF (see attached). I would like to add a check to make sure the idw is the "latest" vaulted file.  Is it possible with iLogic to check the vault file status? Using Inventor 2024. Thanks,Kirk

**Code:**

```vb
Imports Autodesk.DataManagement.Client.Framework.Vault
Imports Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports AWS = Autodesk.Connectivity.WebServices
Imports AWST = Autodesk.Connectivity.WebServicesTools
Imports VB = Connectivity.Application.VaultBase
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency'.Properties
AddReference "Autodesk.DataManagement.Client.Framework.Vault.dll"
AddReference "Autodesk.DataManagement.Client.Framework.dll"
AddReference "Connectivity.Application.VaultBase.dll"
AddReference "Autodesk.Connectivity.WebServices.dll"
Sub Main
	Dim sLocalFile As String = ThisDoc.PathAndFileName(True)
	Logger.Info("sLocalFile = " & sLocalFile)
	Dim bIsLatestVersion As Boolean = LocalFileIsLatestVersion(sLocalFile)
	MsgBox("Specified File Is Latest Version = " & bIsLatestVersion.ToString(), , "")
End Sub

Function LocalFileIsLatestVersion(fullFileName As String) As Boolean
	'validate the input
	If String.IsNullOrWhiteSpace(fullFileName) Then
		Logger.Debug("Empty value passed into the 'LocalFileIsLatestVersion' method!")
		Return False
	End If
	If Not System.IO.File.Exists(fullFileName) Then
		Logger.Debug("File specified as input into the 'LocalFileIsLatestVersion' method does not exist!")
		Return False
	End If

	'get Vault connection
	Dim oConn As VDF.Vault.Currency.Connections.Connection
	oConn = VB.ConnectionManager.Instance.Connection

	'if no connection then exit routine
	If oConn Is Nothing Then
		Logger.Debug("Could not get Vault Connection!")
		MessageBox.Show("Not Logged In to Vault! - Login first and repeat executing this rule.")
		Exit Function
	End If

	'convert the 'local' path to a 'Vault' path
	Dim sVaultPath As String = iLogicVault.ConvertLocalPathToVaultPath(fullFileName)
	Logger.Info("sVaultPath = " & sVaultPath)

	'get the 'Folder' object
	Dim oFolder As AWS.Folder = oConn.WebServiceManager.DocumentService.GetFolderByPath(sVaultPath)
	If oFolder Is Nothing Then
		Logger.Warn("Did NOT get the Vault Folder that this file was in.")
	Else
		Logger.Info("Got the Vault Folder that this file was in.")
	End If
	
	'get a dictionary of all property definitions
	'if we provide an 'EmptyCategory' instead of 'Nothing', we get no entries
	Dim oPropDefsDict As VDF.Vault.Currency.Properties.PropertyDefinitionDictionary
	oPropDefsDict = oConn.PropertyManager.GetPropertyDefinitions( _
	VDF.Vault.Currency.Entities.EntityClassIds.Files, _
	Nothing, _
	VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeAll)
	
	If (oPropDefsDict Is Nothing) OrElse (oPropDefsDict.Count = 0) Then
		Logger.Warn("PropertyDefinitionDictionary was either Nothing or Empty!")
	Else
		Logger.Info("Got the PropertyDefinitionDictionary, and it had " & oPropDefsDict.Count.ToString() & " entries.")
	End If

	'get the VaultStatus PropertyDefinition
	Dim oStatusPropDef As VDF.Vault.Currency.Properties.PropertyDefinition
	oStatusPropDef = oPropDefsDict(VDF.Vault.Currency.Properties.PropertyDefinitionIds.Client.VaultStatus)

	If (oStatusPropDef Is Nothing) Then
		Logger.Warn("PropertyDefinition was Nothing!")
	Else
		Logger.Info("Got the PropertyDefinition OK.")
	End If

	'get all child File objects in that folder, except the 'hidden' ones
	Dim oChildFiles As AWS.File() = oConn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(oFolder.Id, False)
	'if no child files found, then exit routine
	If (oChildFiles Is Nothing) OrElse (oChildFiles.Length = 0) Then
		Logger.Warn("No child files in specified folder!")
		Return False
	Else
		Logger.Info("Found " & oChildFiles.Length.ToString() & " child files in specified folder!")
	End If

	'start iterating through each file in the folder
	For Each oFile As AWS.File In oChildFiles
		'only process one specific file, by its file name
		If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
		Logger.Info("Found the matching Vault file.")
		
		'get the FileIteration object of this File object
		Dim oFileIt As VDF.Vault.Currency.Entities.FileIteration
		oFileIt = New VDF.Vault.Currency.Entities.FileIteration(oConn, oFile)
		If oFileIt IsNot Nothing Then
			Logger.Info("Got the FileIteration for this file.")
		Else
			Logger.Warn("Did NOT get the FileIteration for this file.")
			Continue For
		End If
		
		'Dim oPropExtProv As VDF.Vault.Interfaces.IPropertyExtensionProvider
		'oPropExtProv = oConn.PropertyManager.
		
		'Dim oPropValueSettings As New VDF.Vault.Currency.Properties.PropertyValueSettings()
		'oPropValueSettings.AddPropertyExtensionProvider(oPropExtProv)
				
		'read value of VaultStatus Property of specified File
		Dim oStatus As VDF.Vault.Currency.Properties.EntityStatusImageInfo
		oStatus = oConn.PropertyManager.GetPropertyValue(oFileIt, oStatusPropDef, Nothing)
		
		'check value, and respond accordingly
		If oStatus.Status.VersionState = VDF.Vault.Currency.Properties.EntityStatus.VersionStateEnum.MatchesLatestVaultVersion Then
			Logger.Info("Following File Version Matches Latest Vault Version:" _
			& vbCrLf & oFile.Name)
			Return True
		Else
			Logger.Info("Following File Version Does Not Match Latest Vault Version:" _
			& vbCrLf & oFile.Name)
		End If
	Next
	Return False
End Function
```

```vb
Imports Autodesk.DataManagement.Client.Framework.Vault
Imports Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports AWS = Autodesk.Connectivity.WebServices
Imports AWST = Autodesk.Connectivity.WebServicesTools
Imports VB = Connectivity.Application.VaultBase
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency'.Properties
AddReference "Autodesk.DataManagement.Client.Framework.Vault.dll"
AddReference "Autodesk.DataManagement.Client.Framework.dll"
AddReference "Connectivity.Application.VaultBase.dll"
AddReference "Autodesk.Connectivity.WebServices.dll"
Sub Main
	Dim sLocalFile As String = ThisDoc.PathAndFileName(True)
	Logger.Info("sLocalFile = " & sLocalFile)
	Dim bIsLatestVersion As Boolean = LocalFileIsLatestVersion(sLocalFile)
	MsgBox("Specified File Is Latest Version = " & bIsLatestVersion.ToString(), , "")
End Sub

Function LocalFileIsLatestVersion(fullFileName As String) As Boolean
	'validate the input
	If String.IsNullOrWhiteSpace(fullFileName) Then
		Logger.Debug("Empty value passed into the 'LocalFileIsLatestVersion' method!")
		Return False
	End If
	If Not System.IO.File.Exists(fullFileName) Then
		Logger.Debug("File specified as input into the 'LocalFileIsLatestVersion' method does not exist!")
		Return False
	End If

	'get Vault connection
	Dim oConn As VDF.Vault.Currency.Connections.Connection
	oConn = VB.ConnectionManager.Instance.Connection

	'if no connection then exit routine
	If oConn Is Nothing Then
		Logger.Debug("Could not get Vault Connection!")
		MessageBox.Show("Not Logged In to Vault! - Login first and repeat executing this rule.")
		Exit Function
	End If

	'convert the 'local' path to a 'Vault' path
	Dim sVaultPath As String = iLogicVault.ConvertLocalPathToVaultPath(fullFileName)
	Logger.Info("sVaultPath = " & sVaultPath)

	'get the 'Folder' object
	Dim oFolder As AWS.Folder = oConn.WebServiceManager.DocumentService.GetFolderByPath(sVaultPath)
	If oFolder Is Nothing Then
		Logger.Warn("Did NOT get the Vault Folder that this file was in.")
	Else
		Logger.Info("Got the Vault Folder that this file was in.")
	End If
	
	'get a dictionary of all property definitions
	'if we provide an 'EmptyCategory' instead of 'Nothing', we get no entries
	Dim oPropDefsDict As VDF.Vault.Currency.Properties.PropertyDefinitionDictionary
	oPropDefsDict = oConn.PropertyManager.GetPropertyDefinitions( _
	VDF.Vault.Currency.Entities.EntityClassIds.Files, _
	Nothing, _
	VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeAll)
	
	If (oPropDefsDict Is Nothing) OrElse (oPropDefsDict.Count = 0) Then
		Logger.Warn("PropertyDefinitionDictionary was either Nothing or Empty!")
	Else
		Logger.Info("Got the PropertyDefinitionDictionary, and it had " & oPropDefsDict.Count.ToString() & " entries.")
	End If

	'get the VaultStatus PropertyDefinition
	Dim oStatusPropDef As VDF.Vault.Currency.Properties.PropertyDefinition
	oStatusPropDef = oPropDefsDict(VDF.Vault.Currency.Properties.PropertyDefinitionIds.Client.VaultStatus)

	If (oStatusPropDef Is Nothing) Then
		Logger.Warn("PropertyDefinition was Nothing!")
	Else
		Logger.Info("Got the PropertyDefinition OK.")
	End If

	'get all child File objects in that folder, except the 'hidden' ones
	Dim oChildFiles As AWS.File() = oConn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(oFolder.Id, False)
	'if no child files found, then exit routine
	If (oChildFiles Is Nothing) OrElse (oChildFiles.Length = 0) Then
		Logger.Warn("No child files in specified folder!")
		Return False
	Else
		Logger.Info("Found " & oChildFiles.Length.ToString() & " child files in specified folder!")
	End If

	'start iterating through each file in the folder
	For Each oFile As AWS.File In oChildFiles
		'only process one specific file, by its file name
		If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
		Logger.Info("Found the matching Vault file.")
		
		'get the FileIteration object of this File object
		Dim oFileIt As VDF.Vault.Currency.Entities.FileIteration
		oFileIt = New VDF.Vault.Currency.Entities.FileIteration(oConn, oFile)
		If oFileIt IsNot Nothing Then
			Logger.Info("Got the FileIteration for this file.")
		Else
			Logger.Warn("Did NOT get the FileIteration for this file.")
			Continue For
		End If
		
		'Dim oPropExtProv As VDF.Vault.Interfaces.IPropertyExtensionProvider
		'oPropExtProv = oConn.PropertyManager.
		
		'Dim oPropValueSettings As New VDF.Vault.Currency.Properties.PropertyValueSettings()
		'oPropValueSettings.AddPropertyExtensionProvider(oPropExtProv)
				
		'read value of VaultStatus Property of specified File
		Dim oStatus As VDF.Vault.Currency.Properties.EntityStatusImageInfo
		oStatus = oConn.PropertyManager.GetPropertyValue(oFileIt, oStatusPropDef, Nothing)
		
		'check value, and respond accordingly
		If oStatus.Status.VersionState = VDF.Vault.Currency.Properties.EntityStatus.VersionStateEnum.MatchesLatestVaultVersion Then
			Logger.Info("Following File Version Matches Latest Vault Version:" _
			& vbCrLf & oFile.Name)
			Return True
		Else
			Logger.Info("Following File Version Does Not Match Latest Vault Version:" _
			& vbCrLf & oFile.Name)
		End If
	Next
	Return False
End Function
```

```vb
If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
```

```vb
If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
```

```vb
If Not oFile.Name.Equals(System.IO.Path.GetFileName(fullFileName), StringComparison.CurrentCultureIgnoreCase) Then Continue For
```

```vb
If Not oFile.Name.Equals(System.IO.Path.GetFileName(fullFileName), StringComparison.CurrentCultureIgnoreCase) Then Continue For
```

---

## Show Thumbnail in the Windows.Forms.PictureBox iLogic Inventor 2026

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/show-thumbnail-in-the-windows-forms-picturebox-ilogic-inventor/td-p/13760959#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/show-thumbnail-in-the-windows-forms-picturebox-ilogic-inventor/td-p/13760959#messageview_0)

**Author:** pavol_krasnansky1

**Date:** ‎08-09-2025
	
		
		06:47 AM

**Description:** Hi, the iLogic rule below works in Inventor 2023 and does not work in Inventor 2026. Why? What should I change? Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic

AddReference "System.Drawing.dll"
AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"

'------------------------------------------------------------------...

**Code:**

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic

AddReference "System.Drawing.dll"
AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"

'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = Compatibility.VB6.IPictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub
```

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic

AddReference "System.Drawing.dll"
AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"

'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = Compatibility.VB6.IPictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub
```

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic
AddReference "System.Drawing.dll"
'AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"


'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = PictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub

Public Shared Function PictureDispToImage(pictureDisp As stdole.IPictureDisp) As System.Drawing.Image
	Dim oImage As System.Drawing.Image
	If pictureDisp IsNot Nothing Then
		If pictureDisp.Type = 1 Then
			Dim hpalette As IntPtr = New IntPtr(pictureDisp.hPal)
			oImage = oImage.FromHbitmap(New IntPtr(pictureDisp.Handle), hpalette)
		End If
		If pictureDisp.Type = 2 Then
			oImage = New System.Drawing.Imaging.Metafile(New IntPtr(pictureDisp.Handle), New System.Drawing.Imaging.WmfPlaceableFileHeader())
		End If
	End If
	Return oImage
End Function
```

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic
AddReference "System.Drawing.dll"
'AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"


'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = PictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub

Public Shared Function PictureDispToImage(pictureDisp As stdole.IPictureDisp) As System.Drawing.Image
	Dim oImage As System.Drawing.Image
	If pictureDisp IsNot Nothing Then
		If pictureDisp.Type = 1 Then
			Dim hpalette As IntPtr = New IntPtr(pictureDisp.hPal)
			oImage = oImage.FromHbitmap(New IntPtr(pictureDisp.Handle), hpalette)
		End If
		If pictureDisp.Type = 2 Then
			oImage = New System.Drawing.Imaging.Metafile(New IntPtr(pictureDisp.Handle), New System.Drawing.Imaging.WmfPlaceableFileHeader())
		End If
	End If
	Return oImage
End Function
```

---

## Using Ilogic to update model states

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/using-ilogic-to-update-model-states/td-p/12844817](https://forums.autodesk.com/t5/inventor-programming-forum/using-ilogic-to-update-model-states/td-p/12844817)

**Author:** aberningR89Y3

**Date:** ‎06-17-2024
	
		
		10:52 AM

**Description:** I am looking for a way to update every model state in a part file without having to manually select each model state and click 'local update'.Ideally, I would like to do this without having to know the name of every model state beforehand. Something like: find each model state name, put it into an array, and then loop through and useThisApplication.ActiveView.Update()I am not well versed on VBA/iLogic syntax so maybe there is a 'For Each' loop that can do it?   Thanks in advance,Alex
					
				
...

**Code:**

```vb
ThisApplication.ActiveView.Update()
```

```vb
Dim oDoc As Inventor.Document = ThisDoc.FactoryDocument
If oDoc Is Nothing Then Return
Dim oMSs As ModelStates = oDoc.ComponentDefinition.ModelStates
If oMSs Is Nothing OrElse oMSs.Count = 0 Then Return
Dim oAMS As ModelState = oMSs.ActiveModelState
For Each oMS As ModelState In oMSs
	oMS.Activate
	Dim oMSDoc As Inventor.Document = TryCast(oMS.Document, Inventor.Document)
	If oMSDoc Is Nothing Then Continue For
	oMSDoc.Update2(True)
Next oMS
If oMSs.ActiveModelState IsNot oAMS Then oAMS.Activate
```

```vb
Dim oDoc As Inventor.Document = ThisDoc.FactoryDocument
If oDoc Is Nothing Then Return
Dim oMSs As ModelStates = oDoc.ComponentDefinition.ModelStates
If oMSs Is Nothing OrElse oMSs.Count = 0 Then Return
Dim oAMS As ModelState = oMSs.ActiveModelState
For Each oMS As ModelState In oMSs
	oMS.Activate
	Dim oMSDoc As Inventor.Document = TryCast(oMS.Document, Inventor.Document)
	If oMSDoc Is Nothing Then Continue For
	oMSDoc.Update2(True)
Next oMS
If oMSs.ActiveModelState IsNot oAMS Then oAMS.Activate
```

---

## How to &quot;Show All&quot; in a part document with iLogic?

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-show-all-quot-in-a-part-document-with-ilogic/td-p/13832701](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-show-all-quot-in-a-part-document-with-ilogic/td-p/13832701)

**Author:** acheungR3ZK5

**Date:** ‎10-01-2025
	
		
		04:46 AM

**Description:** How do I call the "Show All" command (shown in screenshot) in a part document with iLogic? I know how to loop through all the parts and make them visible if not visible, but it takes longer if there are many parts and you can't undo the whole command in one undo.  
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim partDoc As PartDocument
partDoc = ThisApplication.ActiveDocument

Dim compDef As PartComponentDefinition
compDef = partDoc.ComponentDefinition

Dim activeViewRep As DesignViewRepresentation = compDef.RepresentationsManager.ActiveDesignViewRepresentation
activeViewRep.ShowAll
```

```vb
Dim partDoc As PartDocument
partDoc = ThisApplication.ActiveDocument

Dim compDef As PartComponentDefinition
compDef = partDoc.ComponentDefinition

Dim activeViewRep As DesignViewRepresentation = compDef.RepresentationsManager.ActiveDesignViewRepresentation
activeViewRep.ShowAll
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("ShowAllBodiesCtxCmd").Execute()
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("ShowAllBodiesCtxCmd").Execute()
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("HideOtherBodiesCtxCmd").Execute()
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("HideOtherBodiesCtxCmd").Execute()
```

```vb
Dim oApp As Inventor.Application = ThisApplication
Dim oPDoc As PartDocument = oApp.ActiveDocument
Dim oSBColl As ObjectCollection = oApp.TransientObjects.CreateObjectCollection()
For Each oSB As SurfaceBody In oPDoc.ComponentDefinition.SurfaceBodies
    oSBColl.Add(oSB)
Next
oPDoc.SelectSet.Clear()
oPDoc.SelectSet.SelectMultiple(oSBColl)
oApp.CommandManager.ControlDefinitions.Item("PartVisibilityCtxCmd").Execute()
```

```vb
Dim oApp As Inventor.Application = ThisApplication
Dim oPDoc As PartDocument = oApp.ActiveDocument
Dim oSBColl As ObjectCollection = oApp.TransientObjects.CreateObjectCollection()
For Each oSB As SurfaceBody In oPDoc.ComponentDefinition.SurfaceBodies
    oSBColl.Add(oSB)
Next
oPDoc.SelectSet.Clear()
oPDoc.SelectSet.SelectMultiple(oSBColl)
oApp.CommandManager.ControlDefinitions.Item("PartVisibilityCtxCmd").Execute()
```

---

## Check if drawing parts list is splitted

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/check-if-drawing-parts-list-is-splitted/td-p/12568593#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/check-if-drawing-parts-list-is-splitted/td-p/12568593#messageview_0)

**Author:** goran_nilssonK2TWB

**Date:** ‎02-19-2024
	
		
		12:06 AM

**Description:** Hello! I want to check if drawing parts list is splitted. If it's splitted I want to get a message.My code is using On error resume next, so it's not possible to use a Try.I can't find any information if it's possible or not. /Goran
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Public Sub Main()
	Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If oDDoc Is Nothing Then Exit Sub  
	Dim oBrowserPane As BrowserPane = oDDoc.BrowserPanes.ActivePane
	Dim oTopBrowserNode As BrowserNode = oBrowserPane.TopNode   
	Dim oBrowserNode As BrowserNode    
	For Each oBrowserNode In oTopBrowserNode.BrowserNodes
		Dim oBrowserNode2 As BrowserNode
		For Each oBrowserNode2 In oBrowserNode.BrowserNodes
			Dim oNativeObject As Object = oBrowserNode2.NativeObject
			If TypeName(oNativeObject) = "PartsList" Then
				If oBrowserNode2.BrowserNodes.Count > 0 Then
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split")
				End If
				Exit For
			End If
		Next
	Next       
End Sub
```

```vb
Public Sub Main()
	Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If oDDoc Is Nothing Then Exit Sub  
	Dim oBrowserPane As BrowserPane = oDDoc.BrowserPanes.ActivePane
	Dim oTopBrowserNode As BrowserNode = oBrowserPane.TopNode   
	Dim oBrowserNode As BrowserNode    
	For Each oBrowserNode In oTopBrowserNode.BrowserNodes
		Dim oBrowserNode2 As BrowserNode
		For Each oBrowserNode2 In oBrowserNode.BrowserNodes
			Dim oNativeObject As Object = oBrowserNode2.NativeObject
			If TypeName(oNativeObject) = "PartsList" Then
				If oBrowserNode2.BrowserNodes.Count > 0 Then
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split")
				End If
				Exit For
			End If
		Next
	Next       
End Sub
```

```vb
Sub main
	Dim DrawingDoc = ThisApplication.ActiveEditDocument
	Dim modelState As Integer = 0
	For Each oSheet In DrawingDoc.Sheets
		itemNumber=1
		While itemNumber < 10
			Try
				oPartsList = oSheet.PartsLists.Item(itemNumber)
				If modelState = oPartsList.ReferencedDocumentDescriptor.ReferencedModelState
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split or this is the first occurance of it")
					modelState=oPartsList.ReferencedDocumentDescriptor.ReferencedModelState
				End If
					itemNumber = itemNumber + 1
			Catch
				Exit While
			End Try
		End While
	Next
End Sub
```

---

## Custom ribbon button doesn't execute VBA sub

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-button-doesn-t-execute-vba-sub/td-p/13788549](https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-button-doesn-t-execute-vba-sub/td-p/13788549)

**Author:** sv.lucht

**Date:** ‎08-29-2025
	
		
		04:54 AM

**Description:** I try to create custom ribbon tabs that get called when different criteria are met. The VBA code resides in StampMaker.ivb library in the UIManager module and everything works fine. I can create and destroy tabs,panel,buttons, but I cannot make the buttons work. The functions get never called. I tried to simplify the code, as seen below. I tried to get help from Gemini. But nothing works. So I would be very grateful for any suggestions. Here is a simplified version of the code just for troublesh...

**Code:**

```vb
Public Sub RunCleanButtonTest()
' --- Phase 1: Aggressive Cleanup ---
'Const TEST_MACRO_NAME As String = "Default.Module1.testio"
Const TEST_MACRO_NAME As String = "StampMaker.UIManager.testprompt"
Const TEST_TAB_NAME As String = "MyTestTab"
Const TEST_PANEL_NAME As String = "MyTestPanel"

' Delete the Control Definition
On Error Resume Next
ThisApplication.CommandManager.ControlDefinitions.item(TEST_MACRO_NAME).delete
Debug.Print "Old Control Definition deleted (if it existed)."
On Error GoTo 0

' Delete the Ribbon Tab
Dim zeroRibbon As Inventor.Ribbon
Set zeroRibbon = ThisApplication.UserInterfaceManager.Ribbons.item("ZeroDoc")
On Error Resume Next
zeroRibbon.RibbonTabs.item(TEST_TAB_NAME).delete
Debug.Print "Old Ribbon Tab deleted (if it existed)."
On Error GoTo 0

' --- Phase 2: Create a single, simple button ---
Debug.Print "Creating new UI..."

' Create Tab and Panel
Dim newTab As RibbonTab
Set newTab = zeroRibbon.RibbonTabs.Add("My Test", TEST_TAB_NAME, "ClientID_TestTab")
Dim newPanel As RibbonPanel
Set newPanel = newTab.RibbonPanels.Add("My Panel", TEST_PANEL_NAME, "ClientID_TestPanel")

' Create Button Definition
Dim buttonDef As ButtonDefinition
Set buttonDef = ThisApplication.CommandManager.ControlDefinitions.AddButtonDefinition( _
"Click Me", _
TEST_MACRO_NAME, _
CommandTypesEnum.kNonShapeEditCmdType)

' Add button to panel
Call newPanel.CommandControls.AddButton(buttonDef, True)

Debug.Print "--- Test UI created. Please click the 'Click Me' button on the 'My Test' tab. ---"
End Sub


Public Sub testprompt()
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
End Sub
```

```vb
Public Sub RunCleanButtonTest()
' --- Phase 1: Aggressive Cleanup ---
'Const TEST_MACRO_NAME As String = "Default.Module1.testio"
Const TEST_MACRO_NAME As String = "StampMaker.UIManager.testprompt"
Const TEST_TAB_NAME As String = "MyTestTab"
Const TEST_PANEL_NAME As String = "MyTestPanel"

' Delete the Control Definition
On Error Resume Next
ThisApplication.CommandManager.ControlDefinitions.item(TEST_MACRO_NAME).delete
Debug.Print "Old Control Definition deleted (if it existed)."
On Error GoTo 0

' Delete the Ribbon Tab
Dim zeroRibbon As Inventor.Ribbon
Set zeroRibbon = ThisApplication.UserInterfaceManager.Ribbons.item("ZeroDoc")
On Error Resume Next
zeroRibbon.RibbonTabs.item(TEST_TAB_NAME).delete
Debug.Print "Old Ribbon Tab deleted (if it existed)."
On Error GoTo 0

' --- Phase 2: Create a single, simple button ---
Debug.Print "Creating new UI..."

' Create Tab and Panel
Dim newTab As RibbonTab
Set newTab = zeroRibbon.RibbonTabs.Add("My Test", TEST_TAB_NAME, "ClientID_TestTab")
Dim newPanel As RibbonPanel
Set newPanel = newTab.RibbonPanels.Add("My Panel", TEST_PANEL_NAME, "ClientID_TestPanel")

' Create Button Definition
Dim buttonDef As ButtonDefinition
Set buttonDef = ThisApplication.CommandManager.ControlDefinitions.AddButtonDefinition( _
"Click Me", _
TEST_MACRO_NAME, _
CommandTypesEnum.kNonShapeEditCmdType)

' Add button to panel
Call newPanel.CommandControls.AddButton(buttonDef, True)

Debug.Print "--- Test UI created. Please click the 'Click Me' button on the 'My Test' tab. ---"
End Sub


Public Sub testprompt()
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
End Sub
```

```vb
Public Sub CreateMinimalMacroButton()
' --- 1. Define the names for the macro and UI elements ---
Const MACRO_FULL_NAME As String = "Module1.MyTargetSubroutine"
Const TAB_ID As String = "StampMaker_MinimalTab_1"
Const PANEL_ID As String = "StampMaker_MinimalPanel_1"

Dim uiMgr As UserInterfaceManager
Set uiMgr = ThisApplication.UserInterfaceManager

' --- 2. Clean up any old versions of these UI elements ---
' Using "On Error Resume Next" is a simple way to delete items
' without causing an error if they don't exist yet.
On Error Resume Next

' Delete the Ribbon Tab (this also removes the panel and the button on it)
uiMgr.Ribbons.item("ZeroDoc").RibbonTabs.item(TAB_ID).delete

' Delete the Control Definition (the "brain" of the button)
ThisApplication.CommandManager.ControlDefinitions.item(MACRO_FULL_NAME).delete

' Return to normal error handling
On Error GoTo 0

' --- 3. Create the new UI ---

' Get the ribbon that's visible when no document is open
Dim zeroRibbon As Ribbon
Set zeroRibbon = uiMgr.Ribbons.item("ZeroDoc")

' Create the new tab
Dim myTab As RibbonTab
Set myTab = zeroRibbon.RibbonTabs.Add("Minimal Test", TAB_ID, "ClientID_MinimalTab")

' Create a panel on the tab
Dim myPanel As RibbonPanel
Set myPanel = myTab.RibbonPanels.Add("Test Panel", PANEL_ID, "ClientID_MinimalPanel")

' Create the Macro Definition - this links the button to your code
Dim macroDef As MacroControlDefinition
Set macroDef = ThisApplication.CommandManager.ControlDefinitions.AddMacroControlDefinition(MACRO_FULL_NAME)

' Add the actual button to the panel
Call myPanel.CommandControls.AddMacro(macroDef)

' --- 4. Notify the user ---
MsgBox "Minimal macro button has been created." & vbCrLf & _
"Look for the 'Minimal Test' tab in your ribbon.", vbInformation

End Sub
```

```vb
Public Sub CreateMinimalMacroButton()
' --- 1. Define the names for the macro and UI elements ---
Const MACRO_FULL_NAME As String = "Module1.MyTargetSubroutine"
Const TAB_ID As String = "StampMaker_MinimalTab_1"
Const PANEL_ID As String = "StampMaker_MinimalPanel_1"

Dim uiMgr As UserInterfaceManager
Set uiMgr = ThisApplication.UserInterfaceManager

' --- 2. Clean up any old versions of these UI elements ---
' Using "On Error Resume Next" is a simple way to delete items
' without causing an error if they don't exist yet.
On Error Resume Next

' Delete the Ribbon Tab (this also removes the panel and the button on it)
uiMgr.Ribbons.item("ZeroDoc").RibbonTabs.item(TAB_ID).delete

' Delete the Control Definition (the "brain" of the button)
ThisApplication.CommandManager.ControlDefinitions.item(MACRO_FULL_NAME).delete

' Return to normal error handling
On Error GoTo 0

' --- 3. Create the new UI ---

' Get the ribbon that's visible when no document is open
Dim zeroRibbon As Ribbon
Set zeroRibbon = uiMgr.Ribbons.item("ZeroDoc")

' Create the new tab
Dim myTab As RibbonTab
Set myTab = zeroRibbon.RibbonTabs.Add("Minimal Test", TAB_ID, "ClientID_MinimalTab")

' Create a panel on the tab
Dim myPanel As RibbonPanel
Set myPanel = myTab.RibbonPanels.Add("Test Panel", PANEL_ID, "ClientID_MinimalPanel")

' Create the Macro Definition - this links the button to your code
Dim macroDef As MacroControlDefinition
Set macroDef = ThisApplication.CommandManager.ControlDefinitions.AddMacroControlDefinition(MACRO_FULL_NAME)

' Add the actual button to the panel
Call myPanel.CommandControls.AddMacro(macroDef)

' --- 4. Notify the user ---
MsgBox "Minimal macro button has been created." & vbCrLf & _
"Look for the 'Minimal Test' tab in your ribbon.", vbInformation

End Sub
```

---

## Looking for some &quot;fun&quot; iLogic code for a T-Shirt

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/looking-for-some-quot-fun-quot-ilogic-code-for-a-t-shirt/td-p/13840538#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/looking-for-some-quot-fun-quot-ilogic-code-for-a-t-shirt/td-p/13840538#messageview_0)

**Author:** chris

**Date:** ‎10-06-2025
	
		
		07:08 PM

**Description:** I'm looking to add some fun or unique iLogic code for a T-Shirt that I'm making for the designers at work. Can anyone think of some fun code that might look cool or fun on the back of a shirt? (it can be short or long), just looking for something fun/cool, I plan on using the code edit colors as wellExample: iLogic code to create a cube or sphere, I though about some iProperties, but I'm open to ideas!

**Code:**

```vb
Dim Stress As Double = 0
While Project.Deadline < Now
    Stress += 1
    Sleep(1)
End While
```

```vb
Do Until Sketch.Perfect = True
    IterateDesign()
Loop
```

```vb
Try
    OpenInventor()
    StartDesigning("GeniusPart")
    While Not Done
        Iterate()
        Overthink()
    End While
Catch ex As BurnoutException
    Coffee.Refill()
    Resume Next
End Try
```

```vb
If Coffee.Level < 10 Then
    Call Coffee.Refill()
    Return
End If
Call Design.Part("Brilliance")
```

```vb
If You.AreConstrainedTo(Me) Then
    Life.IsFullyDefined = True
End If
```

```vb
If You.Fit(My.Mate) Then
    Let'sAssemble()
End If
```

```vb
If Not Coffee.Exists Then
    System.Collapse()
End If
```

```vb
Sub EngineeringProcess()
    ' Engineering is spelled: C-H-A-N-G-E
    
    Design = CreateInitialConcept()
    SendToClient()
    
    While (ProjectNotCancelled)
        changeRequest = Client.SendEmail()
        
        Select Case changeRequest.Severity
            Case "Minor tweak"
                RedoEverything()
            Case "Small adjustment"
                StartFromScratch()
            Case "Just one little thing"
                QuestionCareerPath()
        End Select
        
        UpdateRevisionNumber() ' Now at Rev ZZ.47
        
        If (design.LooksLikeOriginalConcept = False) Then
            Console.WriteLine("Nailed it! ��")
        End If
        
        WaitForNextChangeRequest(milliseconds:=3)
    End While
    
    ' Note: ProjectNotCancelled always returns True
End Sub
```

---

## Trailing Zeros

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427#messageview_0)

**Author:** cgchad

**Date:** ‎09-09-2025
	
		
		10:09 AM

**Description:** I know this has been asked dozens of times but for some reason I can't get it to work out.  I am running Inventor Professional 2024.I have 2 parameters that i need to display with just 2 decimal places.  Even if those decimal places are zeros.  No matter what I have done it stops when it's a whole number or has something like .70 it will change to .7. My code for this is currently:'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)
'set ra...

**Code:**

```vb
'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)
'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & WIDTH & " x " & LENGTH
```

```vb
Dim Height As Double = tank_id
Dim Width As Double = tank_height

Dim HeightString As String = CStr(Height)
Dim WidthString As String = CStr(Width)

Dim HeightParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_id")
Dim WidthParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_height")

HeightParam.Precision = 3 'I wish this worked to control the decimals displayed from either the expression or value. 
WidthParam.Precision = 3

Dim HeightSubString As String() = HeightParam.Expression.Split(" "c) 
Dim WidthSubString As String() = WidthParam.Expression.Split(" "c)


Logger.Info("Height blue value: " & tank_id & " Height double: " & Height & " Height String: " & HeightString & " Height param expression: " & HeightParam.Expression & " Height param value direct: " & HeightParam.Value / 2.54 & " Height substring: " & HeightSubString(0))
Logger.Info("Width blue value: "  & tank_height & " Widdth double: " & Width & "  Width String: " & WidthString & " Width param expression: " & WidthParam.Expression & " Width param value direct: " & HeightParam.Value / 2.54 & " Width substring: " & WidthSubString(0))
```

```vb
Dim Height As Double = tank_id
Dim Width As Double = tank_height

Dim HeightString As String = CStr(Height)
Dim WidthString As String = CStr(Width)

Dim HeightParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_id")
Dim WidthParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_height")

HeightParam.Precision = 3 'I wish this worked to control the decimals displayed from either the expression or value. 
WidthParam.Precision = 3

Dim HeightSubString As String() = HeightParam.Expression.Split(" "c) 
Dim WidthSubString As String() = WidthParam.Expression.Split(" "c)


Logger.Info("Height blue value: " & tank_id & " Height double: " & Height & " Height String: " & HeightString & " Height param expression: " & HeightParam.Expression & " Height param value direct: " & HeightParam.Value / 2.54 & " Height substring: " & HeightSubString(0))
Logger.Info("Width blue value: "  & tank_height & " Widdth double: " & Width & "  Width String: " & WidthString & " Width param expression: " & WidthParam.Expression & " Width param value direct: " & HeightParam.Value / 2.54 & " Width substring: " & WidthSubString(0))
```

```vb
'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)
```

```vb
Dim DoubleLength As Double = Round(SheetMetal.FlatExtentsLength, 2)
Dim DoubleWidth As Double = Round(SheetMetal.FlatExtentsWidth, 2)
LENGTH = DoubleLength
WIDTH = DoubleWidth

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & DoubleWidth.ToString("F2") & " x " & DoubleLength.ToString("F2")
```

```vb
Dim DoubleLength As Double = Round(SheetMetal.FlatExtentsLength, 2)
Dim DoubleWidth As Double = Round(SheetMetal.FlatExtentsWidth, 2)
LENGTH = DoubleLength
WIDTH = DoubleWidth

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & DoubleWidth.ToString("F2") & " x " & DoubleLength.ToString("F2")
```

---

## iLogic Drawing Layer Visibility Toggle

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-drawing-layer-visibility-toggle/td-p/13816579](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-drawing-layer-visibility-toggle/td-p/13816579)

**Author:** dave_novak

**Date:** ‎09-18-2025
	
		
		10:09 AM

**Description:** Hello, I would like to figure out (with a little help from my friends) how to get Drawings to show the correct Visibility State in the Layer Drop Down after running the following iLogic Rule: If ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True
	ElseIf ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True Then
	ThisDrawing.Document.StylesManager.Lay...

**Code:**

```vb
If ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True
	ElseIf ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False
	
End If

iLogicVb.UpdateWhenDone = True
```

---

## Inventor DWG to AutoCAD DWG Conversion

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/inventor-dwg-to-autocad-dwg-conversion/td-p/11339593#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/inventor-dwg-to-autocad-dwg-conversion/td-p/11339593#messageview_0)

**Author:** Harshil.JAINS8RVT

**Date:** ‎08-04-2022
	
		
		02:29 AM

**Description:** Hello, I have a requirement to convert from Inventor DWG to AutoCAD DWG. I have tried the code as shown below:- ApplicationAddInsPtr pAddIns = pApplication->GetApplicationAddIns();TranslatorAddInPtr pTrans;_bstr_t ClientId = _T("{C24E3AC2-122E-11D5-8E91-0010B541CD80}");pTrans = pAddIns->GetItemById(ClientId);DocumentPtr pINVDoc = pApplication->GetActiveDocument();short bHasSaveOption = pTrans->GetHasSaveCopyAsOptions(pINVDoc, oContext, oOptions);if (bHasSaveOption){_bstr_t bstrAttribValue("D:\\C...

**Code:**

```vb
class ThisRule { 
public:
    CComPtr<Application> ThisApplication;

    void Main() {

        HRESULT Result = NOERROR;

        CComPtr<Document> gDoc;
        Result = ThisApplication->get_ActiveDocument(&gDoc);

        CComPtr<DrawingDocument> doc;
        doc = gDoc;
        
        CComPtr<TranslationContext> context;
        Result = ThisApplication->TransientObjects->CreateTranslationContext(&context);
        context->put_Type(kFileBrowseIOMechanism);

        CComPtr<NameValueMap> options;
        Result = ThisApplication->TransientObjects->CreateNameValueMap(&options);

        CComPtr<DataMedium> dataMedium;
        Result = ThisApplication->TransientObjects->CreateDataMedium(&dataMedium);

        CComPtr<TranslatorAddIn> translator;
        translator = GetDwgTranslator();


        if (translator->HasSaveCopyAsOptions[doc][context][options]) {
            _bstr_t iniFileName = _T("D:\\forum\\\\dwgExport.ini");

            VARIANT varProtType;
            varProtType.vt = VT_BSTR;
            varProtType.bstrVal = iniFileName;

            Result = options->put_Value(_T("Export_Acad_IniFile"), varProtType);
        }

        _bstr_t dwgFileName("D:\\forum\\Drawing1_AutoCAD.dwg");
        Result = dataMedium->put_FileName(dwgFileName);

        Result = translator->SaveCopyAs(doc, context, options, dataMedium);
    }

private:
    CComPtr<TranslatorAddIn> GetDwgTranslator() {
        HRESULT Result = NOERROR;

        _bstr_t ClientId = _T("{C24E3AC2-122E-11D5-8E91-0010B541CD80}");

        CComPtr<ApplicationAddIn> addin;
        Result = ThisApplication->ApplicationAddIns->get_ItemById(ClientId, &addin);

        CComPtr<TranslatorAddIn> translator;
        translator = addin;

        return translator;
    }
};
```

```vb
class ThisRule { 
public:
    CComPtr<Application> ThisApplication;

    void Main() {

        HRESULT Result = NOERROR;

        CComPtr<Document> gDoc;
        Result = ThisApplication->get_ActiveDocument(&gDoc);

        CComPtr<DrawingDocument> doc;
        doc = gDoc;
        
        CComPtr<TranslationContext> context;
        Result = ThisApplication->TransientObjects->CreateTranslationContext(&context);
        context->put_Type(kFileBrowseIOMechanism);

        CComPtr<NameValueMap> options;
        Result = ThisApplication->TransientObjects->CreateNameValueMap(&options);

        CComPtr<DataMedium> dataMedium;
        Result = ThisApplication->TransientObjects->CreateDataMedium(&dataMedium);

        CComPtr<TranslatorAddIn> translator;
        translator = GetDwgTranslator();


        if (translator->HasSaveCopyAsOptions[doc][context][options]) {
            _bstr_t iniFileName = _T("D:\\forum\\\\dwgExport.ini");

            VARIANT varProtType;
            varProtType.vt = VT_BSTR;
            varProtType.bstrVal = iniFileName;

            Result = options->put_Value(_T("Export_Acad_IniFile"), varProtType);
        }

        _bstr_t dwgFileName("D:\\forum\\Drawing1_AutoCAD.dwg");
        Result = dataMedium->put_FileName(dwgFileName);

        Result = translator->SaveCopyAs(doc, context, options, dataMedium);
    }

private:
    CComPtr<TranslatorAddIn> GetDwgTranslator() {
        HRESULT Result = NOERROR;

        _bstr_t ClientId = _T("{C24E3AC2-122E-11D5-8E91-0010B541CD80}");

        CComPtr<ApplicationAddIn> addin;
        Result = ThisApplication->ApplicationAddIns->get_ItemById(ClientId, &addin);

        CComPtr<TranslatorAddIn> translator;
        translator = addin;

        return translator;
    }
};
```

```vb
ThisRule rule;
rule.ThisApplication = GetInventor();
rule.Main();
```

```vb
ThisRule rule;
rule.ThisApplication = GetInventor();
rule.Main();
```

```vb
Sub Main()
    If Not ConfirmProceed() Then Return

    Dim sourceFolder As String = GetFolder("Select folder containing DWG files")
    If sourceFolder = "" Then Return

    Dim exportFolder As String = GetFolder("Select folder to save AutoCAD DWG files")
    If exportFolder = "" Then Return

    Dim configPath As String = "C:\Users\thoma\Documents\$_INDUSTRIAL_DRAFTING_&_DESIGN\INVENTOR_FILES\PRESETS\EXPORT-DWG.ini"
    If Not System.IO.File.Exists(configPath) Then
        MessageBox.Show("Missing DWG config: " & configPath)
        Return
    End If

    Dim files() As String = System.IO.Directory.GetFiles(sourceFolder, "*.dwg")
    Dim successCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim errorLog As New System.Text.StringBuilder
    Dim exportLog As New System.Text.StringBuilder

    exportLog.AppendLine("DWG Export Log - " & Now.ToString())
    exportLog.AppendLine("")

    For Each file As String In files
        Dim doc As Document = Nothing
        Try
            doc = ThisApplication.Documents.Open(file, False)
            If doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                doc.Close(True)
                Continue For
            End If

            Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(file)
            Dim drawDoc As DrawingDocument = doc
            Dim sheet1 As Sheet = drawDoc.Sheets(1)
            If sheet1.DrawingViews.Count = 0 Then Throw New Exception("No views on Sheet 1")
            Dim baseView As DrawingView = sheet1.DrawingViews(1)

            Dim modelDoc As Document = baseView.ReferencedDocumentDescriptor.ReferencedDocument
            If modelDoc Is Nothing Then Throw New Exception("Referenced model is missing.")

            Dim modelRev As String = ""
            Try
                modelRev = iProperties.Value(modelDoc, "Project", "Revision Number")
            Catch
                modelRev = "NO-REV"
            End Try

            Dim exportName As String = baseName & "-REV-" & SanitizeFileName(modelRev) & ".dwg"
            Dim exportPath As String = System.IO.Path.Combine(exportFolder, exportName)

            exportLog.AppendLine("DWG: " & baseName)
            exportLog.AppendLine(" ↳ Model: " & modelDoc.DisplayName)
            exportLog.AppendLine(" ↳ REV: " & modelRev)

            If System.IO.File.Exists(exportPath) Then
                exportLog.AppendLine(" ⏭ Skipped (same REV already exported)")
                doc.Close(True)
                skippedCount += 1
                Continue For
            End If

            ' Set up translator
            Dim addIn As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
            Dim context = ThisApplication.TransientObjects.CreateTranslationContext
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim options = ThisApplication.TransientObjects.CreateNameValueMap
            If addIn.HasSaveCopyAsOptions(doc, context, options) Then
                options.Value("Export_Acad_IniFile") = configPath
            End If
            Dim oData = ThisApplication.TransientObjects.CreateDataMedium
            oData.FileName = exportPath

            addIn.SaveCopyAs(doc, context, options, oData)
            doc.Close(True)
            successCount += 1
            exportLog.AppendLine(" ✅ Exported: " & exportName)
        Catch ex As Exception
            If Not doc Is Nothing Then doc.Close(True)
            errorLog.AppendLine("ERROR processing " & file & ": " & ex.Message)
        End Try

        exportLog.AppendLine("")
    Next

    ' Write logs
    System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "DWG_Export_Log.txt"), exportLog.ToString())
    If errorLog.Length > 0 Then
        System.IO.File.WriteAllText(System.IO.Path.Combine(exportFolder, "DWG_Export_Errors.txt"), errorLog.ToString())
    End If

    MessageBox.Show("DWG export complete." & vbCrLf & _
        "✔ Files exported: " & successCount & vbCrLf & _
        "⏭ Skipped (already exported at this REV): " & skippedCount & vbCrLf & _
        "⚠ Errors: " & errorLog.Length, "Export Summary")
End Sub

Function GetFolder(prompt As String) As String
    Dim dlg As New FolderBrowserDialog
    dlg.Description = prompt
    If dlg.ShowDialog() <> DialogResult.OK Then Return ""
    Return dlg.SelectedPath
End Function

Function ConfirmProceed() As Boolean
    Dim result = MessageBox.Show("Export DWGs only if target REV file does not exist?", "Confirm Export", MessageBoxButtons.YesNo)
    Return result = DialogResult.Yes
End Function

Function SanitizeFileName(input As String) As String
    For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
        input = input.Replace(c, "_")
    Next
    Return input
End Function
```

---

## Can't Debug Inventor 2025 in VS2019

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/can-t-debug-inventor-2025-in-vs2019/td-p/13003292](https://forums.autodesk.com/t5/inventor-programming-forum/can-t-debug-inventor-2025-in-vs2019/td-p/13003292)

**Author:** tdant

**Date:** ‎09-05-2024
	
		
		10:30 AM

**Description:** Hey all,I'm working on debugging some addins that have been functioning fine since Inventor 2021, and have run into a new unfamiliar error since updating to 2025:Apparently this is some kind of memory write access denial, but it might as well be Greek to me. I've enabled native code debugging and both Microsoft Symbol Servers and NuGet.org Symbol Server, but am still getting the error. Any ideas?

**Code:**

```vb
<PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
   <BaseAddress>285212672</BaseAddress>
   <ConfigurationOverrideFile>
   </ConfigurationOverrideFile>
   <DocumentationFile>
   </DocumentationFile>
   <FileAlignment>4096</FileAlignment>
   <NoStdLib>false</NoStdLib>
   <NoWarn>
   </NoWarn>
   <RegisterForComInterop>false</RegisterForComInterop>
   <RemoveIntegerChecks>false</RemoveIntegerChecks>
   <CodeAnalysisRuleSet>AllRules.ruleset</CodeAnalysisRuleSet>
  <OutputPath>C:\Users\MeMySelf\AppData\Roaming\Autodesk\ApplicationPlugins\HKInventorAddin\</OutputPath>
  <AppendTargetFrameworkToOutputPath>false</AppendTargetFrameworkToOutputPath>
 </PropertyGroup>
```

```vb
<PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
   <BaseAddress>285212672</BaseAddress>
   <ConfigurationOverrideFile>
   </ConfigurationOverrideFile>
   <DocumentationFile>
   </DocumentationFile>
   <FileAlignment>4096</FileAlignment>
   <NoStdLib>false</NoStdLib>
   <NoWarn>
   </NoWarn>
   <RegisterForComInterop>false</RegisterForComInterop>
   <RemoveIntegerChecks>false</RemoveIntegerChecks>
   <CodeAnalysisRuleSet>AllRules.ruleset</CodeAnalysisRuleSet>
  <OutputPath>C:\Users\MeMySelf\AppData\Roaming\Autodesk\ApplicationPlugins\HKInventorAddin\</OutputPath>
  <AppendTargetFrameworkToOutputPath>false</AppendTargetFrameworkToOutputPath>
 </PropertyGroup>
```

---

## Can't Debug Inventor 2025 in VS2019

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/can-t-debug-inventor-2025-in-vs2019/td-p/13003292#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/can-t-debug-inventor-2025-in-vs2019/td-p/13003292#messageview_0)

**Author:** tdant

**Date:** ‎09-05-2024
	
		
		10:30 AM

**Description:** Hey all,I'm working on debugging some addins that have been functioning fine since Inventor 2021, and have run into a new unfamiliar error since updating to 2025:Apparently this is some kind of memory write access denial, but it might as well be Greek to me. I've enabled native code debugging and both Microsoft Symbol Servers and NuGet.org Symbol Server, but am still getting the error. Any ideas?

**Code:**

```vb
<PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
   <BaseAddress>285212672</BaseAddress>
   <ConfigurationOverrideFile>
   </ConfigurationOverrideFile>
   <DocumentationFile>
   </DocumentationFile>
   <FileAlignment>4096</FileAlignment>
   <NoStdLib>false</NoStdLib>
   <NoWarn>
   </NoWarn>
   <RegisterForComInterop>false</RegisterForComInterop>
   <RemoveIntegerChecks>false</RemoveIntegerChecks>
   <CodeAnalysisRuleSet>AllRules.ruleset</CodeAnalysisRuleSet>
  <OutputPath>C:\Users\MeMySelf\AppData\Roaming\Autodesk\ApplicationPlugins\HKInventorAddin\</OutputPath>
  <AppendTargetFrameworkToOutputPath>false</AppendTargetFrameworkToOutputPath>
 </PropertyGroup>
```

```vb
<PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
   <BaseAddress>285212672</BaseAddress>
   <ConfigurationOverrideFile>
   </ConfigurationOverrideFile>
   <DocumentationFile>
   </DocumentationFile>
   <FileAlignment>4096</FileAlignment>
   <NoStdLib>false</NoStdLib>
   <NoWarn>
   </NoWarn>
   <RegisterForComInterop>false</RegisterForComInterop>
   <RemoveIntegerChecks>false</RemoveIntegerChecks>
   <CodeAnalysisRuleSet>AllRules.ruleset</CodeAnalysisRuleSet>
  <OutputPath>C:\Users\MeMySelf\AppData\Roaming\Autodesk\ApplicationPlugins\HKInventorAddin\</OutputPath>
  <AppendTargetFrameworkToOutputPath>false</AppendTargetFrameworkToOutputPath>
 </PropertyGroup>
```

---

## Code to measure length and width

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/code-to-measure-length-and-width/td-p/13792602](https://forums.autodesk.com/t5/inventor-programming-forum/code-to-measure-length-and-width/td-p/13792602)

**Author:** berry.lejeune

**Date:** ‎09-02-2025
	
		
		03:23 AM

**Description:** Hi all, Don't even know if it's possible but I'll ask anywayWe've received a Revit model of a big building which is covered in stone plates. Now I need to get those dimensions of those plates for our manufacturer.But in stead of measuring each indivudial one, is it possible with a code to click on the surface and that it gives me the length and the width? Thanks

**Code:**

```vb
'Dimensions

Geo_Length = Measure.PreciseExtentsLength
Geo_Width = Measure.PreciseExtentsWidth
Geo_Height = Measure.PreciseExtentsHeight
```

```vb
Dim oInvApp As Inventor.Application = ThisApplication
Dim oCM As CommandManager = oInvApp.CommandManager
Dim oUOM As UnitsOfMeasure = oInvApp.ActiveDocument.UnitsOfMeasure
'Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kFootLengthUnits
Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kCentimeterLengthUnits

Do
	Dim oFace As Object
	oFace = oCM.Pick(SelectionFilterEnum.kPartFaceFilter, "Select Face...")
	If oFace Is Nothing Then Exit Do
	Dim oBox As Box2d = oFace.Evaluator.ParamRangeRect
	Dim oSizes(1) As Double
	oSizes(0) = Round(oUOM.ConvertUnits(oBox.MaxPoint.X - oBox.MinPoint.X,
										eLeng, oUOM.LengthUnits), 3)
	oSizes(1) = Round(oUOM.ConvertUnits(oBox.MaxPoint.Y - oBox.MinPoint.Y,
										eLeng, oUOM.LengthUnits), 3)
	MessageBox.Show("Length - " & Abs(oSizes.Max) & vbLf &
					"Width - " & Abs(oSizes.Min), "Surface size:")
	Logger.Info("Length - " & Abs(oSizes.Max))
	Logger.Info("Width - " & Abs(oSizes.Min))
Loop
```

```vb
Dim oInvApp As Inventor.Application = ThisApplication
Dim oCM As CommandManager = oInvApp.CommandManager
Dim oUOM As UnitsOfMeasure = oInvApp.ActiveDocument.UnitsOfMeasure
'Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kFootLengthUnits
Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kCentimeterLengthUnits

Do
	Dim oFace As Object
	oFace = oCM.Pick(SelectionFilterEnum.kPartFaceFilter, "Select Face...")
	If oFace Is Nothing Then Exit Do
	Dim oBox As Box2d = oFace.Evaluator.ParamRangeRect
	Dim oSizes(1) As Double
	oSizes(0) = Round(oUOM.ConvertUnits(oBox.MaxPoint.X - oBox.MinPoint.X,
										eLeng, oUOM.LengthUnits), 3)
	oSizes(1) = Round(oUOM.ConvertUnits(oBox.MaxPoint.Y - oBox.MinPoint.Y,
										eLeng, oUOM.LengthUnits), 3)
	MessageBox.Show("Length - " & Abs(oSizes.Max) & vbLf &
					"Width - " & Abs(oSizes.Min), "Surface size:")
	Logger.Info("Length - " & Abs(oSizes.Max))
	Logger.Info("Width - " & Abs(oSizes.Min))
Loop
```

---

## Stop an iLogic rule while it's running

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/stop-an-ilogic-rule-while-it-s-running/td-p/12304171#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/stop-an-ilogic-rule-while-it-s-running/td-p/12304171#messageview_0)

**Author:** Daan_M

**Date:** ‎10-13-2023
	
		
		06:31 AM

**Description:** Hi, I'm currently running a loop on 250+ files taking about 2hrs, i forgot one line of code. Is it possible to cancel your rule during a run?  if so, how do i do this?
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
For Each ...
	If IO.File.exists("C:\Temp\IF_YOU_DELETE_THIS_CODE_WILL_STOP.txt")
                MsgBox("rule " & iLogicVb.RuleName & " stopped because text file doesn't exist")                'logger.info("rule " & iLogicVb.RuleName & " stopped because text file doesn't exist")                return                 'exit sub                'goto ...
	End If
	...
Next
```

```vb
Public Sub Main
    progBar = New TestProgressBar(ThisApplication)
    progBar.Start()
End Sub

Public Class TestProgressBar
    Private WithEvents progBar As Inventor.ProgressBar
    Private invApp As Inventor.Application

    Public Sub New(InventorApp As Inventor.Application)
        invApp = InventorApp
    End Sub

    Public Sub Start()
        progBar = invApp.CreateProgressBar(False, 10, "Test of Progress Bar", True)

        Dim j As Integer
        For j = 0 To 10
            progBar.Message = ("Current Index: " & j & "/" & 10)
            Threading.Thread.Sleep(1000)
            progBar.UpdateProgress()
        Next

        progBar.Close()
    End Sub

    Private Sub progBar_OnCancel() Handles progBar.OnCancel
        progBar.Close()
        MsgBox("Cancelled")
    End Sub
End Class
```

```vb
' iLogic: Save all open documents, skipping any that error.
' Drops dialogs (SilentOperation) and restores settings afterward.

Dim app As Inventor.Application = ThisApplication
Dim docs As Inventor.Documents = app.Documents

Dim total As Integer = 0
Dim saved As Integer = 0
Dim failed As New System.Collections.Generic.List(Of String)

' Remember app settings
Dim originalSilent As Boolean = app.SilentOperation
Dim originalScreen As Boolean = app.ScreenUpdating

Try
    app.SilentOperation = True
    app.ScreenUpdating = False

    ' Take a snapshot of currently open docs (collection can change while saving)
    Dim openDocs As New System.Collections.Generic.List(Of Inventor.Document)
    For Each d As Inventor.Document In docs
        openDocs.Add(d)
    Next

    For Each d As Inventor.Document In openDocs
        total += 1
        Try
            d.Save()                           ' Try to save everything
            saved += 1
        Catch ex As System.Exception
            failed.Add(d.DisplayName & " — " & ex.Message)
            ' Skip and continue with the next document
        End Try
    Next

Finally
    app.SilentOperation = originalSilent
    app.ScreenUpdating = originalScreen
End Try

Dim report As String = _
    "Save All — iLogic" & vbCrLf & vbCrLf & _
    "Open docs: " & total & vbCrLf & _
    "Saved: " & saved & vbCrLf & _
    "Failed: " & failed.Count

If failed.Count > 0 Then
    report &= vbCrLf & vbCrLf & "Failures:" & vbCrLf & "— " & String.Join(vbCrLf & "— ", failed.ToArray())
End If

MsgBox(report, vbInformation, "iLogic Save All")
```

```vb
' iLogic: Save all open documents, skipping any that error.
' Drops dialogs (SilentOperation) and restores settings afterward.

Dim app As Inventor.Application = ThisApplication
Dim docs As Inventor.Documents = app.Documents

Dim total As Integer = 0
Dim saved As Integer = 0
Dim failed As New System.Collections.Generic.List(Of String)

' Remember app settings
Dim originalSilent As Boolean = app.SilentOperation
Dim originalScreen As Boolean = app.ScreenUpdating

Try
    app.SilentOperation = True
    app.ScreenUpdating = False

    ' Take a snapshot of currently open docs (collection can change while saving)
    Dim openDocs As New System.Collections.Generic.List(Of Inventor.Document)
    For Each d As Inventor.Document In docs
        openDocs.Add(d)
    Next

    For Each d As Inventor.Document In openDocs
        total += 1
        Try
            d.Save()                           ' Try to save everything
            saved += 1
        Catch ex As System.Exception
            failed.Add(d.DisplayName & " — " & ex.Message)
            ' Skip and continue with the next document
        End Try
    Next

Finally
    app.SilentOperation = originalSilent
    app.ScreenUpdating = originalScreen
End Try

Dim report As String = _
    "Save All — iLogic" & vbCrLf & vbCrLf & _
    "Open docs: " & total & vbCrLf & _
    "Saved: " & saved & vbCrLf & _
    "Failed: " & failed.Count

If failed.Count > 0 Then
    report &= vbCrLf & vbCrLf & "Failures:" & vbCrLf & "— " & String.Join(vbCrLf & "— ", failed.ToArray())
End If

MsgBox(report, vbInformation, "iLogic Save All")
```

---

## Use filename of 2D file, to fill in iproperties - Inventor 2024

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/use-filename-of-2d-file-to-fill-in-iproperties-inventor-2024/td-p/13802821](https://forums.autodesk.com/t5/inventor-programming-forum/use-filename-of-2d-file-to-fill-in-iproperties-inventor-2024/td-p/13802821)

**Author:** jens_herrebaut83

**Date:** ‎09-09-2025
	
		
		04:46 AM

**Description:** Hi,
 
Is there a way to save a 2D file (ex: Title-desciption-customer.idw) and it automatically fills in the iproperties of that 2D file?
 
 
@jens_herrebaut83 Your post title was modified to add the product name and version and to increase findability - CGBenner
 
 
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Try
	Dim oDoc As Document = ThisApplication.ActiveEditDocument
	Dim oName As String = ThisDoc.FileName(False)
	Dim oFirstDashPos As Integer = InStr(oName, "-")
	Dim oSecondDashPos As Integer = InStr(oFirstDashPos + 1, oName, "-")
	Dim oTitle As String = Left(oName, oFirstDashPos -1)
	Dim oClient As String = Mid(oName, oFirstDashPos + 1, oSecondDashPos - oFirstDashPos - 1)
	Dim oDesc As String = Right(oName, Len(oName) - oSecondDashPos)
	
	iProperties.Value("Summary", "Title") = oTitle
	iProperties.Value("Custom", "Client") = oClient
	iProperties.Value("Project", "Description") = oDesc
Catch
End Try
```

```vb
Try
	Dim oDoc As Document = ThisApplication.ActiveEditDocument
	Dim oName As String = ThisDoc.FileName(False)
	Dim oFirstDashPos As Integer = InStr(oName, "-")
	Dim oSecondDashPos As Integer = InStr(oFirstDashPos + 1, oName, "-")
	Dim oTitle As String = Left(oName, oFirstDashPos -1)
	Dim oClient As String = Mid(oName, oFirstDashPos + 1, oSecondDashPos - oFirstDashPos - 1)
	Dim oDesc As String = Right(oName, Len(oName) - oSecondDashPos)
	
	iProperties.Value("Summary", "Title") = oTitle
	iProperties.Value("Custom", "Client") = oClient
	iProperties.Value("Project", "Description") = oDesc
Catch
End Try
```

---

## Workflow to modify MaterialAsset (apply another PhysicalPropertiesAsset)

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/workflow-to-modify-materialasset-apply-another/td-p/13822777#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/workflow-to-modify-materialasset-apply-another/td-p/13822777#messageview_0)

**Author:** Maxim-CADman77

**Date:** ‎09-23-2025
	
		
		04:12 PM

**Description:** I'm studying the possibility to automate modification of MaterialAssets (apply another PhysicalPropertiesAsset according to some logic).PhysicalPropertiesAsset description says it is a writable Asset property and "When assigning physical properties, the physical properties asset must exist in the same document as the material." Ok. I do have two materials target and source in my Part document.I then try to assign PhysicalPropertiesAsset of source material to PhysicalPropertiesAsset of target mat...

**Code:**

```vb
Dim ptDoc As PartDocument = ThisDoc.Document

Logger.Info(ptDoc.MaterialAssets.Count)

Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Logger.Info(srcMat.DisplayName)
Logger.Info(vbTab & srcMat.PhysicalPropertiesAsset.DisplayName)

Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
Logger.Info(tgtMat.DisplayName)
Logger.Info(vbTab & tgtMat.PhysicalPropertiesAsset.DisplayName)

tgtMat.PhysicalPropertiesAsset = srcMat.PhysicalPropertiesAsset ' HRESULT: 0x80020003 (DISP_E_MEMBERNOTFOUND)
```

```vb
Dim ptDoc As PartDocument = ThisDoc.Document

Logger.Info(ptDoc.MaterialAssets.Count)

Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Logger.Info(srcMat.DisplayName)
Logger.Info(vbTab & srcMat.PhysicalPropertiesAsset.DisplayName)

Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
Logger.Info(tgtMat.DisplayName)
Logger.Info(vbTab & tgtMat.PhysicalPropertiesAsset.DisplayName)

tgtMat.PhysicalPropertiesAsset = srcMat.PhysicalPropertiesAsset ' HRESULT: 0x80020003 (DISP_E_MEMBERNOTFOUND)
```

```vb
Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
```

```vb
Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
```

```vb
Dim srcMat As MaterialAsset = ptDoc.MaterialAssets(1)
Dim tgtMat As MaterialAsset = ptDoc.MaterialAssets(2)
```

```vb
Dim srcMat As MaterialAsset = ptDoc.MaterialAssets(1)
Dim tgtMat As MaterialAsset = ptDoc.MaterialAssets(2)
```

---

## Check Vault Status

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/check-vault-status/td-p/13834312](https://forums.autodesk.com/t5/inventor-programming-forum/check-vault-status/td-p/13834312)

**Author:** karthur1

**Date:** ‎10-02-2025
	
		
		05:27 AM

**Description:** I have an iLogic rule to quickly print to PDF (see attached). I would like to add a check to make sure the idw is the "latest" vaulted file.  Is it possible with iLogic to check the vault file status? Using Inventor 2024. Thanks,Kirk

**Code:**

```vb
Imports Autodesk.DataManagement.Client.Framework.Vault
Imports Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports AWS = Autodesk.Connectivity.WebServices
Imports AWST = Autodesk.Connectivity.WebServicesTools
Imports VB = Connectivity.Application.VaultBase
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency'.Properties
AddReference "Autodesk.DataManagement.Client.Framework.Vault.dll"
AddReference "Autodesk.DataManagement.Client.Framework.dll"
AddReference "Connectivity.Application.VaultBase.dll"
AddReference "Autodesk.Connectivity.WebServices.dll"
Sub Main
	Dim sLocalFile As String = ThisDoc.PathAndFileName(True)
	Logger.Info("sLocalFile = " & sLocalFile)
	Dim bIsLatestVersion As Boolean = LocalFileIsLatestVersion(sLocalFile)
	MsgBox("Specified File Is Latest Version = " & bIsLatestVersion.ToString(), , "")
End Sub

Function LocalFileIsLatestVersion(fullFileName As String) As Boolean
	'validate the input
	If String.IsNullOrWhiteSpace(fullFileName) Then
		Logger.Debug("Empty value passed into the 'LocalFileIsLatestVersion' method!")
		Return False
	End If
	If Not System.IO.File.Exists(fullFileName) Then
		Logger.Debug("File specified as input into the 'LocalFileIsLatestVersion' method does not exist!")
		Return False
	End If

	'get Vault connection
	Dim oConn As VDF.Vault.Currency.Connections.Connection
	oConn = VB.ConnectionManager.Instance.Connection

	'if no connection then exit routine
	If oConn Is Nothing Then
		Logger.Debug("Could not get Vault Connection!")
		MessageBox.Show("Not Logged In to Vault! - Login first and repeat executing this rule.")
		Exit Function
	End If

	'convert the 'local' path to a 'Vault' path
	Dim sVaultPath As String = iLogicVault.ConvertLocalPathToVaultPath(fullFileName)
	Logger.Info("sVaultPath = " & sVaultPath)

	'get the 'Folder' object
	Dim oFolder As AWS.Folder = oConn.WebServiceManager.DocumentService.GetFolderByPath(sVaultPath)
	If oFolder Is Nothing Then
		Logger.Warn("Did NOT get the Vault Folder that this file was in.")
	Else
		Logger.Info("Got the Vault Folder that this file was in.")
	End If
	
	'get a dictionary of all property definitions
	'if we provide an 'EmptyCategory' instead of 'Nothing', we get no entries
	Dim oPropDefsDict As VDF.Vault.Currency.Properties.PropertyDefinitionDictionary
	oPropDefsDict = oConn.PropertyManager.GetPropertyDefinitions( _
	VDF.Vault.Currency.Entities.EntityClassIds.Files, _
	Nothing, _
	VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeAll)
	
	If (oPropDefsDict Is Nothing) OrElse (oPropDefsDict.Count = 0) Then
		Logger.Warn("PropertyDefinitionDictionary was either Nothing or Empty!")
	Else
		Logger.Info("Got the PropertyDefinitionDictionary, and it had " & oPropDefsDict.Count.ToString() & " entries.")
	End If

	'get the VaultStatus PropertyDefinition
	Dim oStatusPropDef As VDF.Vault.Currency.Properties.PropertyDefinition
	oStatusPropDef = oPropDefsDict(VDF.Vault.Currency.Properties.PropertyDefinitionIds.Client.VaultStatus)

	If (oStatusPropDef Is Nothing) Then
		Logger.Warn("PropertyDefinition was Nothing!")
	Else
		Logger.Info("Got the PropertyDefinition OK.")
	End If

	'get all child File objects in that folder, except the 'hidden' ones
	Dim oChildFiles As AWS.File() = oConn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(oFolder.Id, False)
	'if no child files found, then exit routine
	If (oChildFiles Is Nothing) OrElse (oChildFiles.Length = 0) Then
		Logger.Warn("No child files in specified folder!")
		Return False
	Else
		Logger.Info("Found " & oChildFiles.Length.ToString() & " child files in specified folder!")
	End If

	'start iterating through each file in the folder
	For Each oFile As AWS.File In oChildFiles
		'only process one specific file, by its file name
		If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
		Logger.Info("Found the matching Vault file.")
		
		'get the FileIteration object of this File object
		Dim oFileIt As VDF.Vault.Currency.Entities.FileIteration
		oFileIt = New VDF.Vault.Currency.Entities.FileIteration(oConn, oFile)
		If oFileIt IsNot Nothing Then
			Logger.Info("Got the FileIteration for this file.")
		Else
			Logger.Warn("Did NOT get the FileIteration for this file.")
			Continue For
		End If
		
		'Dim oPropExtProv As VDF.Vault.Interfaces.IPropertyExtensionProvider
		'oPropExtProv = oConn.PropertyManager.
		
		'Dim oPropValueSettings As New VDF.Vault.Currency.Properties.PropertyValueSettings()
		'oPropValueSettings.AddPropertyExtensionProvider(oPropExtProv)
				
		'read value of VaultStatus Property of specified File
		Dim oStatus As VDF.Vault.Currency.Properties.EntityStatusImageInfo
		oStatus = oConn.PropertyManager.GetPropertyValue(oFileIt, oStatusPropDef, Nothing)
		
		'check value, and respond accordingly
		If oStatus.Status.VersionState = VDF.Vault.Currency.Properties.EntityStatus.VersionStateEnum.MatchesLatestVaultVersion Then
			Logger.Info("Following File Version Matches Latest Vault Version:" _
			& vbCrLf & oFile.Name)
			Return True
		Else
			Logger.Info("Following File Version Does Not Match Latest Vault Version:" _
			& vbCrLf & oFile.Name)
		End If
	Next
	Return False
End Function
```

```vb
Imports Autodesk.DataManagement.Client.Framework.Vault
Imports Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports AWS = Autodesk.Connectivity.WebServices
Imports AWST = Autodesk.Connectivity.WebServicesTools
Imports VB = Connectivity.Application.VaultBase
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency'.Properties
AddReference "Autodesk.DataManagement.Client.Framework.Vault.dll"
AddReference "Autodesk.DataManagement.Client.Framework.dll"
AddReference "Connectivity.Application.VaultBase.dll"
AddReference "Autodesk.Connectivity.WebServices.dll"
Sub Main
	Dim sLocalFile As String = ThisDoc.PathAndFileName(True)
	Logger.Info("sLocalFile = " & sLocalFile)
	Dim bIsLatestVersion As Boolean = LocalFileIsLatestVersion(sLocalFile)
	MsgBox("Specified File Is Latest Version = " & bIsLatestVersion.ToString(), , "")
End Sub

Function LocalFileIsLatestVersion(fullFileName As String) As Boolean
	'validate the input
	If String.IsNullOrWhiteSpace(fullFileName) Then
		Logger.Debug("Empty value passed into the 'LocalFileIsLatestVersion' method!")
		Return False
	End If
	If Not System.IO.File.Exists(fullFileName) Then
		Logger.Debug("File specified as input into the 'LocalFileIsLatestVersion' method does not exist!")
		Return False
	End If

	'get Vault connection
	Dim oConn As VDF.Vault.Currency.Connections.Connection
	oConn = VB.ConnectionManager.Instance.Connection

	'if no connection then exit routine
	If oConn Is Nothing Then
		Logger.Debug("Could not get Vault Connection!")
		MessageBox.Show("Not Logged In to Vault! - Login first and repeat executing this rule.")
		Exit Function
	End If

	'convert the 'local' path to a 'Vault' path
	Dim sVaultPath As String = iLogicVault.ConvertLocalPathToVaultPath(fullFileName)
	Logger.Info("sVaultPath = " & sVaultPath)

	'get the 'Folder' object
	Dim oFolder As AWS.Folder = oConn.WebServiceManager.DocumentService.GetFolderByPath(sVaultPath)
	If oFolder Is Nothing Then
		Logger.Warn("Did NOT get the Vault Folder that this file was in.")
	Else
		Logger.Info("Got the Vault Folder that this file was in.")
	End If
	
	'get a dictionary of all property definitions
	'if we provide an 'EmptyCategory' instead of 'Nothing', we get no entries
	Dim oPropDefsDict As VDF.Vault.Currency.Properties.PropertyDefinitionDictionary
	oPropDefsDict = oConn.PropertyManager.GetPropertyDefinitions( _
	VDF.Vault.Currency.Entities.EntityClassIds.Files, _
	Nothing, _
	VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeAll)
	
	If (oPropDefsDict Is Nothing) OrElse (oPropDefsDict.Count = 0) Then
		Logger.Warn("PropertyDefinitionDictionary was either Nothing or Empty!")
	Else
		Logger.Info("Got the PropertyDefinitionDictionary, and it had " & oPropDefsDict.Count.ToString() & " entries.")
	End If

	'get the VaultStatus PropertyDefinition
	Dim oStatusPropDef As VDF.Vault.Currency.Properties.PropertyDefinition
	oStatusPropDef = oPropDefsDict(VDF.Vault.Currency.Properties.PropertyDefinitionIds.Client.VaultStatus)

	If (oStatusPropDef Is Nothing) Then
		Logger.Warn("PropertyDefinition was Nothing!")
	Else
		Logger.Info("Got the PropertyDefinition OK.")
	End If

	'get all child File objects in that folder, except the 'hidden' ones
	Dim oChildFiles As AWS.File() = oConn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(oFolder.Id, False)
	'if no child files found, then exit routine
	If (oChildFiles Is Nothing) OrElse (oChildFiles.Length = 0) Then
		Logger.Warn("No child files in specified folder!")
		Return False
	Else
		Logger.Info("Found " & oChildFiles.Length.ToString() & " child files in specified folder!")
	End If

	'start iterating through each file in the folder
	For Each oFile As AWS.File In oChildFiles
		'only process one specific file, by its file name
		If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
		Logger.Info("Found the matching Vault file.")
		
		'get the FileIteration object of this File object
		Dim oFileIt As VDF.Vault.Currency.Entities.FileIteration
		oFileIt = New VDF.Vault.Currency.Entities.FileIteration(oConn, oFile)
		If oFileIt IsNot Nothing Then
			Logger.Info("Got the FileIteration for this file.")
		Else
			Logger.Warn("Did NOT get the FileIteration for this file.")
			Continue For
		End If
		
		'Dim oPropExtProv As VDF.Vault.Interfaces.IPropertyExtensionProvider
		'oPropExtProv = oConn.PropertyManager.
		
		'Dim oPropValueSettings As New VDF.Vault.Currency.Properties.PropertyValueSettings()
		'oPropValueSettings.AddPropertyExtensionProvider(oPropExtProv)
				
		'read value of VaultStatus Property of specified File
		Dim oStatus As VDF.Vault.Currency.Properties.EntityStatusImageInfo
		oStatus = oConn.PropertyManager.GetPropertyValue(oFileIt, oStatusPropDef, Nothing)
		
		'check value, and respond accordingly
		If oStatus.Status.VersionState = VDF.Vault.Currency.Properties.EntityStatus.VersionStateEnum.MatchesLatestVaultVersion Then
			Logger.Info("Following File Version Matches Latest Vault Version:" _
			& vbCrLf & oFile.Name)
			Return True
		Else
			Logger.Info("Following File Version Does Not Match Latest Vault Version:" _
			& vbCrLf & oFile.Name)
		End If
	Next
	Return False
End Function
```

```vb
If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
```

```vb
If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
```

```vb
If Not oFile.Name.Equals(System.IO.Path.GetFileName(fullFileName), StringComparison.CurrentCultureIgnoreCase) Then Continue For
```

```vb
If Not oFile.Name.Equals(System.IO.Path.GetFileName(fullFileName), StringComparison.CurrentCultureIgnoreCase) Then Continue For
```

---

## Inventor 2026 Inventor Application, GetActiveObject is not a member of Marshal

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2026-inventor-application-getactiveobject-is-not-a/td-p/13830311#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2026-inventor-application-getactiveobject-is-not-a/td-p/13830311#messageview_0)

**Author:** A.Acheson

**Date:** ‎09-29-2025
	
		
		02:31 PM

**Description:** I am having trouble converting a VB.Net form designed to run in the ilogic environment. Error on Line 56 : 'GetActiveObject' is not a member of 'Marshal'. I am not using visual studio so have no access to how the application object is created.  Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	 ...

**Code:**

```vb
Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	    'Note: if you shut down the Inventor session that was started
	    'this(way) there is still an Inventor.exe running. We will use
	    'this Boolean to test whether or not the Inventor App  will
	    'need to be shut down.
	   Catch ex2 As Exception
	     MsgBox(ex2.ToString())
	     MsgBox("Unable to get or start Inventor")
	   End Try
End Try
```

```vb
Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	    'Note: if you shut down the Inventor session that was started
	    'this(way) there is still an Inventor.exe running. We will use
	    'this Boolean to test whether or not the Inventor App  will
	    'need to be shut down.
	   Catch ex2 As Exception
	     MsgBox(ex2.ToString())
	     MsgBox("Unable to get or start Inventor")
	   End Try
End Try
```

```vb
_invApp = Thisapplication
```

```vb
_invApp = Thisapplication
```

```vb
' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
			'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
			'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
            'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub
```

```vb
' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
			'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
			'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
            'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub
```

```vb
_invApp.Visible = True
```

```vb
_invApp.Visible = True
```

```vb
ThisApplication.Visible = True
```

```vb
ThisApplication.Visible = True
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
```

```vb
Public Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	Public thisApp As Inventor.Application
	Public Sub New
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub

		
End Class

Sub Main()
 Try
	thisApp = ThisApplication
	oDrawDoc = ThisApplication.ActiveDocument
	MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
  Catch ex As Exception
  End Try

	Dim oMyForm As New WinForm 
	oMyForm.Show

End Sub
```

```vb
Public Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	Public thisApp As Inventor.Application
	Public Sub New
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub

		
End Class

Sub Main()
 Try
	thisApp = ThisApplication
	oDrawDoc = ThisApplication.ActiveDocument
	MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
  Catch ex As Exception
  End Try

	Dim oMyForm As New WinForm 
	oMyForm.Show

End Sub
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
 

 

 Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	'Public  thisApp As Inventor.Application=ThisApplication
	Public Sub New(thisApp As Inventor.Application)
		oDrawDoc = thisApp.ActiveDocument
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		Me.Close
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub


		
End Class

Sub Main()


	Dim oMyForm As New WinForm (ThisApplication)
	oMyForm.Show

End Sub
```

```vb
Option Explicit On
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
 

 

 Class WinForm
	
	Inherits System.Windows.Forms.Form

	Public oArial10 As System.Drawing.Font = New Font("Arial", 10)
	Public oArial12Bold As System.Drawing.Font = New Font("Arial", 10, FontStyle.Bold)
	Public oArial12BoldUnderLined As System.Drawing.Font = New Font("Arial", 12, FontStyle.Underline)
	Public oDrawDoc As DrawingDocument 
	'Public  thisApp As Inventor.Application=ThisApplication
	Public Sub New(thisApp As Inventor.Application)
		oDrawDoc = thisApp.ActiveDocument
		Dim oForm As System.Windows.Forms.Form = Me
		With oForm
			.FormBorderStyle = FormBorderStyle.FixedToolWindow
			.StartPosition = FormStartPosition.CenterScreen
			
			 'Location = New System.Drawing.Point(0, 0)
			.StartPosition = FormStartPosition.Manual
			.Location = New System.Drawing.Point(oForm.Location.X +oForm.Width*4 , _ 
      										oForm.Location.Y+ oForm.Height*1)		  
			.Width = 450
			.Height = 350
			.TopMost = True
			.Font = oArial10
			.Text = "Revision Changes"
			.Name = "Revision Changes"
			.ShowInTaskbar = False
		End With
	
		'[Rev Selection Type
		Dim oRevSelTypeBoxLabel As New System.Windows.Forms.Label
		With oRevSelTypeBoxLabel
			.Text = "Rev Selection Type"
			.Font = oArial12Bold
			.Top = 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oRevSelTypeBoxLabel)
		Dim oRevSelTypeBox As New ComboBox()
		Dim oRevSelTypeList() As String = {"New Rev", "Update Existing"} 

		With oRevSelTypeBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oRevSelTypeList)
			.SelectedIndex = 1
			.Top = oRevSelTypeBoxLabel.Bottom + 1
			.Left = 25
			.Width = 150
			.Name = "Rev Selection Type"
		End With	
		oForm.Controls.Add(oRevSelTypeBox)
			']
		
		'[Designer
		Dim oDesignerBoxLabel As New System.Windows.Forms.Label
		With oDesignerBoxLabel
			.Text = "Designed By"
			.Font = oArial12Bold
			.Top = oRevSelTypeBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 150
		End With
		oForm.Controls.Add(oDesignerBoxLabel)
		Dim oDesignerBox As New ComboBox()
		Dim oDesignerList() As String = {"AA","BB"}
		With oDesignerBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oDesignerList)
			.SelectedIndex = 0 'Index of default selection
			.Top = oDesignerBoxLabel.Bottom + 1
			.Left = 25
			.Width = 50
			.Name = "DesignerBox"
		End With	
		oForm.Controls.Add(oDesignerBox)
			']
		'[Description
		Dim oDescriptionBoxLabel As New System.Windows.Forms.Label
		With oDescriptionBoxLabel
			.Text = "Description"
			.Font = oArial12Bold
			.Top = oDesignerBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 100
		End With
		oForm.Controls.Add(oDescriptionBoxLabel)
		Dim oDescriptionBox As New ComboBox()
		Dim oDescriptionList() As String = {"RELEASE TO PRODUCTION","BB"}
		With oDescriptionBox
			.DropDownStyle = ComboBoxStyle.DropDown
			.Items.Clear()
			.Items.AddRange(oDescriptionList)
			.SelectedIndex = 1
			.Top = oDescriptionBoxLabel.Bottom + 1
			.Left = 25
			.Width = 300
			.Height = 25
			.Name = "DescriptionBox"
		End With	
		oForm.Controls.Add(oDescriptionBox)
			']
			
		'[Approved
		Dim oApprovedBoxLabel As New System.Windows.Forms.Label
		With oApprovedBoxLabel
			.Text = "Approved By"
			.Font = oArial12Bold
			.Top = oDescriptionBox.Bottom + 10
			.Left = 25
			.Height = 20
			.Width = 160
		End With
		oForm.Controls.Add(oApprovedBoxLabel)
		Dim oApprovedBox As New ComboBox()
		Dim oApprovedList() As String = {"AA","BB","CC","DD"}
		With oApprovedBox
			.DropDownStyle = ComboBoxStyle.DropDownList
			.Items.Clear()
			.Items.AddRange(oApprovedList)
			.SelectedIndex = 1
			.Top = oApprovedBoxLabel.Bottom + 1
			.Left = 25
			.Width = 60
			.Name = "ApprovedBox"
		End With	
		oForm.Controls.Add(oApprovedBox)
		']

		Dim oStartButton As New Button()
		With oStartButton
			.Text = "APPLY"
			.Top = oApprovedBox.Bottom + 10
			.Left = 25
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "StartBox"
		End With
		oForm.AcceptButton = oStartButton
		oForm.Controls.Add(oStartButton)
		AddHandler oStartButton.Click, AddressOf oStartButton_Click
		Dim oCancelButton As New Button
		With oCancelButton
			.Text = "CANCEL"
			.Top = oStartButton.Top
			.Left = oStartButton.Right + 10
			.Height = 25
			.Width = 75
			.Enabled = True
			.Name = "CancelButton"
		End With
		oForm.CancelButton = oCancelButton
		oForm.Controls.Add(oCancelButton)
		AddHandler oCancelButton.Click, AddressOf oCancelButton_Click
	End Sub

	'Populate the values from the form
	Public Sub oStartButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	
		Dim Choice As String = Me.Controls.Item("Rev Selection Type").Text
		Dim oArray2 As String =  Me.Controls.Item("DesignerBox").Text
		Dim oArray4 As String = Me.Controls.Item("DescriptionBox").Text
		Dim oArray5 As String = Me.Controls.Item("ApprovedBox").Text
		MessageBox.Show("Here", "Title")
		MessageBox.Show(oDrawDoc.FullDocumentName, "Title")
		Me.Close
End Sub



Public Sub oCancelButton_Click(ByVal oSender As System.Object, ByVal oEventArgs As System.EventArgs)
	Me.Close
End Sub


		
End Class

Sub Main()


	Dim oMyForm As New WinForm (ThisApplication)
	oMyForm.Show

End Sub
```

---

## problem with rules on remote pc

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/problem-with-rules-on-remote-pc/td-p/13810727#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/problem-with-rules-on-remote-pc/td-p/13810727#messageview_0)

**Author:** Frank_DoucetLYPFL

**Date:** ‎09-15-2025
	
		
		02:08 AM

**Description:** i made a rules in inventor for copying and renaming all files from one folder to another, at home it worked like it schould, at work we use remote pc and there i did not work, getting errorLine29 : 'directory' is not declared, it may be inaccessible due to its protection levelLine30 : 'directory' is not declared, it may be inaccessible due to its protection levelLine34 : 'directory' is not declared, it may be inaccessible due to its protection levelLine40 : 'path' is not declared, it may be inac...

**Code:**

```vb
Imports System.IO

Sub Main()
' Define source folder
Dim sourceFolder As String = "C:\Users\Frank\Documents\Inventor\ilogic Afsluiter"

' Ask user to choose destination folder
Dim fbd As New FolderBrowserDialog()
fbd.Description = "Select the destination folder"
fbd.ShowNewFolderButton = True

Dim result As DialogResult = fbd.ShowDialog()
If result <> DialogResult.OK Then
MessageBox.Show("Operation cancelled.", "iLogic")
Exit Sub
End If

Dim destFolder As String = fbd.SelectedPath

' Ask user for prefix (e.g. M593)
Dim prefix As String = InputBox("Enter the prefix (example: M593)", "File Prefix", "M593")
If prefix = "" Then
MessageBox.Show("No prefix entered. Operation cancelled.", "iLogic")
Exit Sub
End If

' Create destination folder if it doesn’t exist
If Not Directory.Exists(destFolder) Then
Directory.CreateDirectory(destFolder)
End If

' Get all files in source folder
Dim files() As String = Directory.GetFiles(sourceFolder)

' Counter for renaming
Dim counter As Integer = 0

For Each filePath As String In files
Dim ext As String = Path.GetExtension(filePath)
Dim newName As String = prefix & "-" & counter.ToString("000") & "-00" & ext
Dim destPath As String = Path.Combine(destFolder, newName)

' Copy file (overwrite if already exists)
File.Copy(filePath, destPath, True)

counter += 1
Next

MessageBox.Show("Files copied and renamed successfully to:" & vbCrLf & destFolder, "iLogic")
End Sub
```

```vb
Imports System.IO

Sub Main()
' Define source folder
Dim sourceFolder As String = "C:\Users\Frank\Documents\Inventor\ilogic Afsluiter"

' Ask user to choose destination folder
Dim fbd As New FolderBrowserDialog()
fbd.Description = "Select the destination folder"
fbd.ShowNewFolderButton = True

Dim result As DialogResult = fbd.ShowDialog()
If result <> DialogResult.OK Then
MessageBox.Show("Operation cancelled.", "iLogic")
Exit Sub
End If

Dim destFolder As String = fbd.SelectedPath

' Ask user for prefix (e.g. M593)
Dim prefix As String = InputBox("Enter the prefix (example: M593)", "File Prefix", "M593")
If prefix = "" Then
MessageBox.Show("No prefix entered. Operation cancelled.", "iLogic")
Exit Sub
End If

' Create destination folder if it doesn’t exist
If Not Directory.Exists(destFolder) Then
Directory.CreateDirectory(destFolder)
End If

' Get all files in source folder
Dim files() As String = Directory.GetFiles(sourceFolder)

' Counter for renaming
Dim counter As Integer = 0

For Each filePath As String In files
Dim ext As String = Path.GetExtension(filePath)
Dim newName As String = prefix & "-" & counter.ToString("000") & "-00" & ext
Dim destPath As String = Path.Combine(destFolder, newName)

' Copy file (overwrite if already exists)
File.Copy(filePath, destPath, True)

counter += 1
Next

MessageBox.Show("Files copied and renamed successfully to:" & vbCrLf & destFolder, "iLogic")
End Sub
```

---

## Load second vba project file or reference subroutine in other project file

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/load-second-vba-project-file-or-reference-subroutine-in-other/td-p/4861128](https://forums.autodesk.com/t5/inventor-programming-forum/load-second-vba-project-file-or-reference-subroutine-in-other/td-p/4861128)

**Author:** pball

**Date:** ‎03-06-2014
	
		
		08:45 AM

**Description:** I'm looking to use two vba project files. I'll have my main ivb file and there is a second ivb file with a single script in it. I do not want to copy that script to the other project file. Is there a way to load or link the second project file from the main one?Another thing is I've manually loaded a second project file for testing and I can run scripts from it via the Macros option on the Tools ribbon bar but I cannot add them to the ribbon bar under the customize user commands option. If there...

**Code:**

```vb
Public Sub LoadVBAProject() LoadAnotherVBAProject "Put Location of VBA project here"End SubPublic Sub LoadAnotherVBAProject(pvarLocation As String) ' counter to work out the new project number Dim varVBACounter As Long varVBACounter = ThisApplication.VBAProjects.Count ' Open the other project ThisApplication.VBAProjects.Open pvarLocation ' get a reference to the just opened VBA project Dim varTempProject As InventorVBAProject Set varTempProject = ThisApplication.VBAProjects(varVBACounter + 1) ' references for all the different parts of the VBA project Dim varVBAComponent As InventorVBAComponent Dim varVBAMember As InventorVBAMember Dim varVBAArgument As InventorVBAArgument ' go through all the components of the Project and display them For Each varVBAComponent In varTempProject.InventorVBAComponents ' print to the immediate window the name of the component Debug.Print varVBAComponent.Name For Each varVBAMember In varVBAComponent.InventorVBAMembers ' print to the immediate window the name of the members of the component Debug.Print " " & varVBAMember.Name For Each varVBAArgument In varVBAMember.Arguments Debug.Print " " & varVBAArgument.Name & "----" & varVBAArgument.ArgumentType Next varVBAArgument Next varVBAMember Next varVBAComponentEnd Sub
```

---

## Export Drawing dimension to excel

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062](https://forums.autodesk.com/t5/inventor-programming-forum/export-drawing-dimension-to-excel/td-p/10155062)

**Author:** eladm

**Date:** ‎03-14-2021
	
		
		12:02 AM

**Description:** Hi I need help with ilogic rule to export all drawing dimensions to ExcelNot VBA , some of the dimeson  have toleranceregards
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i=i+1
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	'MsgBox (b.Tolerance.Upper)
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	'MsgBox (b.Tolerance.Upper)
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	
	GoExcel.CellValue("A" & 1) = "Dimension"
	GoExcel.CellValue("B" & 1) = "Tolerance Upper"
	GoExcel.CellValue("C" & 1) = "Tolerance Lower"
	GoExcel.Save

	Next
```

```vb
Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
	
	GoExcel.CellValue("C:\TEMP\DrawingDimensions.xlsx", "Sheet1", "A" & i) = b.Text.Text.ToString
	GoExcel.CellValue("B" & i) = b.Tolerance.Upper
	GoExcel.CellValue("C" & i) = b.Tolerance.Lower
	
	GoExcel.CellValue("A" & 1) = "Dimension"
	GoExcel.CellValue("B" & 1) = "Tolerance Upper"
	GoExcel.CellValue("C" & 1) = "Tolerance Lower"
	GoExcel.Save

	Next
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper
oCells.item(i,3).value  = b.Tolerance.Lower
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper
oCells.item(i,3).value  = b.Tolerance.Lower
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

oCells.item(2,1).value= "sjgjsagjlk"'b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet

Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet

Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'	'MsgBox (b.Tolerance.Upper)
	
oCells.item(i,1).value = b.Text.Text.ToString
oCells.item(i,2).value = b.Tolerance.Upper*10
oCells.item(i,3).value  = b.Tolerance.Lower*10
	
oCells.item(1,1).value  = "Dimension"
oCells.item(1,2).value = "Tolerance Upper"
oCells.item(1,3).value = "Tolerance Lower"
'	GoExcel.Save()

	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add'.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

'oCells.item(2,1).value= b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'MsgBox (b.Text.Text & " " & b.Tolerance.ToleranceType)
	
	If b.Tolerance.ToleranceType = 31236 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31233 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31235 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = "+/-" & b.Tolerance.Upper*10
		'oCells.item(i,3).value  = "+/-" &b.Tolerance.Lower*10.ToString
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If


	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

```vb
AddReference "Microsoft.Office.Interop.Excel.dll"
Imports exe = Microsoft.Office.Interop.Excel

'Dim oXLFile As String = "C:\TEMP\DrawingDimensions.xlsx"
Dim oSheet As String = "Sheet1"

Dim oExcel As New Microsoft.Office.Interop.Excel.ApplicationClass
Dim oWB As Microsoft.Office.Interop.Excel.Workbook = oExcel.Workbooks.add'.Open(oXLFile)
Dim oWS As Microsoft.Office.Interop.Excel.Worksheet = oWB.Worksheets.Item(oSheet)

Dim oCells As Microsoft.Office.Interop.Excel.Range = oWS.Cells

'oCells.item(2,1).value= b.Text.Text.ToString

Dim a As Inventor.DrawingDocument = ThisDrawing.Document
Dim s As Inventor.Sheet = a.ActiveSheet


Dim i As Integer = 1
For Each b As Inventor.DrawingDimension In s.DrawingDimensions
	i = i + 1
'MsgBox (b.Text.Text & " " & b.Tolerance.ToleranceType)
	
	If b.Tolerance.ToleranceType = 31236 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31233 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = b.Tolerance.Upper*10
		oCells.item(i,3).value  = b.Tolerance.Lower*10
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If
If b.Tolerance.ToleranceType = 31235 Then
		oCells.item(i,1).value = b.Text.Text
		oCells.item(i,2).value = "+/-" & b.Tolerance.Upper*10
		'oCells.item(i,3).value  = "+/-" &b.Tolerance.Lower*10.ToString
	
		oCells.item(1,1).value  = "Dimension"
		oCells.item(1,2).value = "Tolerance Upper"
		oCells.item(1,3).value = "Tolerance Lower"
		End If


	Next
	oWB.saveas("C:\TEMP\"& a.DisplayName & ".xlsx")
oWB.close
```

---

## iLogic Code to export Sheets to IDWS

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-export-sheets-to-idws/td-p/13705767#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-export-sheets-to-idws/td-p/13705767#messageview_0)

**Author:** rossPJ9W7

**Date:** ‎07-01-2025
	
		
		05:12 AM

**Description:** I need help with this code. The goal is to export all the sheets in a idw file to seperate idws with the sheet name as the file name. The code works. All the sheets are saved in a new folder created called " Sheets". All new idw files are named with the sheet name. Issue: The new idws dont open. From my understanding the way the code needs to work is, add a number on the sheet names to be a counter. Then deleted all sheets except one and save as new idw file. Then go to the next counter number a...

**Code:**

```vb
Sub Main
	'get the Inventor Application to a variable
	Dim oInvApp As Inventor.Application = ThisApplication
	'get the main Documents object of the application, for use later
	Dim oDocs As Inventor.Documents = oInvApp.Documents
	'get the 'active' Document, and try to 'Cast' it to the DrawingDocument Type
	'if this fails, then no value will be assigned to the variable
	'but no error will happen either
	Dim oDDoc As DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	'check if the variable got assigned a value...if not, then exit rule
	If oDDoc Is Nothing Then Return
	'get the main Sheets collection of the 'active' drawing
	Dim oSheets As Inventor.Sheets = oDDoc.Sheets
	'if only 1 sheet, then exit rule, because noting to do
	If oSheets.Count = 1 Then Return
	'record currently 'active' Sheet, to reset to active at the end
	Dim oASheet As Inventor.Sheet = oDDoc.ActiveSheet
	'get the full path of the active drawing, without final directory separator character
	Dim sPath As String = System.IO.Path.GetDirectoryName(oDDoc.FullFileName)
	'specify name of sub folder to put new drawing documents into for each sheet
	Dim sSubFolder As String = "Sheets"
	'combine main Path with sub folder name, to make new path
	'adds the directory separator character in for us, automatically
	Dim sNewPath As String = System.IO.Path.Combine(sPath, sSubFolder)
	'create the sub directory, if it does not already exist
	If Not System.IO.Directory.Exists(sNewPath) Then
		System.IO.Directory.CreateDirectory(sNewPath)
	End If
	'not using 'For Each' on purpose, to maintain proper order 
	For iSheetIndex As Integer = 1 To oSheets.Count
		'get the Sheet at this index
		Dim oSheet As Inventor.Sheet = oSheets.Item(iSheetIndex)
		'make this sheet the active sheet
		oSheet.Activate()
		Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
		'save copy of this drawing to new file, in background (new copy not open yet)
		oDDoc.SaveAs(sNewDrawingFFN, True)
		'open that new drawing file (invisibly), so we can modify it
		Dim oNewDDoc As DrawingDocument = oDocs.Open(sNewDrawingFFN, False)
		'get Sheets collection of new drawing to a variable
		Dim oNewDocSheets As Inventor.Sheets = oNewDDoc.Sheets
		'get the Sheet at same Index in new drawing (the one to keep)
		Dim oThisNewDocSheet As Inventor.Sheet = oNewDocSheets.Item(iSheetIndex)
		'delete all other sheets in new drawing besides this sheet
		For Each oNewDocSheet As Inventor.Sheet In oNewDocSheets
			If Not oNewDocSheet Is oThisNewDocSheet Then
				Try
					'Optional 'RetainDependentViews' is False by default
					oNewDocSheet.Delete() 
				Catch
					Logger.Error("Error deleting Sheet named:  " & oNewDocSheet.Name)
				End Try
			End If
		Next
		'save new drawing document again, after changes
		oNewDDoc.Save()
		'next line only valid when not visibly opened, and not being referenced
		oNewDDoc.ReleaseReference()
		'if visibly opened, then use oNewDDoc.Close() method instead
	Next 'oSheet
	'activate originally active sheet again
	oASheet.Activate()
	'only clear out all Documents that are 'unreferenced' in Inventor's memory
	'used after ReleaseReference, to clean-up
	oDocs.CloseAll(True)
End Sub
```

```vb
Sub Main
	'get the Inventor Application to a variable
	Dim oInvApp As Inventor.Application = ThisApplication
	'get the main Documents object of the application, for use later
	Dim oDocs As Inventor.Documents = oInvApp.Documents
	'get the 'active' Document, and try to 'Cast' it to the DrawingDocument Type
	'if this fails, then no value will be assigned to the variable
	'but no error will happen either
	Dim oDDoc As DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	'check if the variable got assigned a value...if not, then exit rule
	If oDDoc Is Nothing Then Return
	'get the main Sheets collection of the 'active' drawing
	Dim oSheets As Inventor.Sheets = oDDoc.Sheets
	'if only 1 sheet, then exit rule, because noting to do
	If oSheets.Count = 1 Then Return
	'record currently 'active' Sheet, to reset to active at the end
	Dim oASheet As Inventor.Sheet = oDDoc.ActiveSheet
	'get the full path of the active drawing, without final directory separator character
	Dim sPath As String = System.IO.Path.GetDirectoryName(oDDoc.FullFileName)
	'specify name of sub folder to put new drawing documents into for each sheet
	Dim sSubFolder As String = "Sheets"
	'combine main Path with sub folder name, to make new path
	'adds the directory separator character in for us, automatically
	Dim sNewPath As String = System.IO.Path.Combine(sPath, sSubFolder)
	'create the sub directory, if it does not already exist
	If Not System.IO.Directory.Exists(sNewPath) Then
		System.IO.Directory.CreateDirectory(sNewPath)
	End If
	'not using 'For Each' on purpose, to maintain proper order 
	For iSheetIndex As Integer = 1 To oSheets.Count
		'get the Sheet at this index
		Dim oSheet As Inventor.Sheet = oSheets.Item(iSheetIndex)
		'make this sheet the active sheet
		oSheet.Activate()
		Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
		'save copy of this drawing to new file, in background (new copy not open yet)
		oDDoc.SaveAs(sNewDrawingFFN, True)
		'open that new drawing file (invisibly), so we can modify it
		Dim oNewDDoc As DrawingDocument = oDocs.Open(sNewDrawingFFN, False)
		'get Sheets collection of new drawing to a variable
		Dim oNewDocSheets As Inventor.Sheets = oNewDDoc.Sheets
		'get the Sheet at same Index in new drawing (the one to keep)
		Dim oThisNewDocSheet As Inventor.Sheet = oNewDocSheets.Item(iSheetIndex)
		'delete all other sheets in new drawing besides this sheet
		For Each oNewDocSheet As Inventor.Sheet In oNewDocSheets
			If Not oNewDocSheet Is oThisNewDocSheet Then
				Try
					'Optional 'RetainDependentViews' is False by default
					oNewDocSheet.Delete() 
				Catch
					Logger.Error("Error deleting Sheet named:  " & oNewDocSheet.Name)
				End Try
			End If
		Next
		'save new drawing document again, after changes
		oNewDDoc.Save()
		'next line only valid when not visibly opened, and not being referenced
		oNewDDoc.ReleaseReference()
		'if visibly opened, then use oNewDDoc.Close() method instead
	Next 'oSheet
	'activate originally active sheet again
	oASheet.Activate()
	'only clear out all Documents that are 'unreferenced' in Inventor's memory
	'used after ReleaseReference, to clean-up
	oDocs.CloseAll(True)
End Sub
```

```vb
iSheetIndex.ToString()
```

```vb
Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
```

```vb
Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
```

```vb
'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
```

```vb
'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
```

```vb
Sub Main
	'get the Inventor Application to a variable
	Dim oInvApp As Inventor.Application = ThisApplication
	'get the main Documents object of the application, for use later
	Dim oDocs As Inventor.Documents = oInvApp.Documents
	'get the 'active' Document, and try to 'Cast' it to the DrawingDocument Type
	'if this fails, then no value will be assigned to the variable
	'but no error will happen either
	Dim oDDoc As DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	'check if the variable got assigned a value...if not, then exit rule
	If oDDoc Is Nothing Then Return
	'get the main Sheets collection of the 'active' drawing
	Dim oSheets As Inventor.Sheets = oDDoc.Sheets
	'if only 1 sheet, then exit rule, because noting to do
	If oSheets.Count = 1 Then Return
	'record currently 'active' Sheet, to reset to active at the end
	Dim oASheet As Inventor.Sheet = oDDoc.ActiveSheet
	'get the full path of the active drawing, without final directory separator character
	Dim sPath As String = System.IO.Path.GetDirectoryName(oDDoc.FullFileName)
	'specify name of sub folder to put new drawing documents into for each sheet
	Dim sSubFolder As String = "Sheets"
	'combine main Path with sub folder name, to make new path
	'adds the directory separator character in for us, automatically
	Dim sNewPath As String = System.IO.Path.Combine(sPath, sSubFolder)
	'create the sub directory, if it does not already exist
	If Not System.IO.Directory.Exists(sNewPath) Then
		System.IO.Directory.CreateDirectory(sNewPath)
	End If
	'not using 'For Each' on purpose, to maintain proper order 
	For iSheetIndex As Integer = 1 To oSheets.Count
		'get the Sheet at this index
		Dim oSheet As Inventor.Sheet = oSheets.Item(iSheetIndex)
		'make this sheet the active sheet
		oSheet.Activate()
		Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		'sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")
		'save copy of this drawing to new file, in background (new copy not open yet)
		oDDoc.SaveAs(sNewDrawingFFN , True)
		'open that new drawing file (invisibly), so we can modify it
		Dim oNewDDoc As DrawingDocument = oDocs.Open(sNewDrawingFFN, False)
		'get Sheets collection of new drawing to a variable
		Dim oNewDocSheets As Inventor.Sheets = oNewDDoc.Sheets
		'get the Sheet at same Index in new drawing (the one to keep)
		Dim oThisNewDocSheet As Inventor.Sheet = oNewDocSheets.Item(iSheetIndex)
		'delete all other sheets in new drawing besides this sheet
		For Each oNewDocSheet As Inventor.Sheet In oNewDocSheets
			If Not oNewDocSheet Is oThisNewDocSheet Then
				Try
					'Optional 'RetainDependentViews' is False by default
					oNewDocSheet.Delete() 
				Catch
					Logger.Error("Error deleting Sheet named:  " & oNewDocSheet.Name)
				End Try
			End If
		Next
		'save new drawing document again, after changes
		oNewDDoc.Save()
		'next line only valid when not visibly opened, and not being referenced
		oNewDDoc.ReleaseReference()
		'if visibly opened, then use oNewDDoc.Close() method instead
	Next 'oSheet
	'activate originally active sheet again
	oASheet.Activate()
	'only clear out all Documents that are 'unreferenced' in Inventor's memory
	'used after ReleaseReference, to clean-up
	oDocs.CloseAll(True)
End Sub
```

---

## Show Thumbnail in the Windows.Forms.PictureBox iLogic Inventor 2026

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/show-thumbnail-in-the-windows-forms-picturebox-ilogic-inventor/td-p/13760959](https://forums.autodesk.com/t5/inventor-programming-forum/show-thumbnail-in-the-windows-forms-picturebox-ilogic-inventor/td-p/13760959)

**Author:** pavol_krasnansky1

**Date:** ‎08-09-2025
	
		
		06:47 AM

**Description:** Hi, the iLogic rule below works in Inventor 2023 and does not work in Inventor 2026. Why? What should I change? Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic

AddReference "System.Drawing.dll"
AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"

'------------------------------------------------------------------...

**Code:**

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic

AddReference "System.Drawing.dll"
AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"

'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = Compatibility.VB6.IPictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub
```

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic

AddReference "System.Drawing.dll"
AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"

'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = Compatibility.VB6.IPictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub
```

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic
AddReference "System.Drawing.dll"
'AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"


'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = PictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub

Public Shared Function PictureDispToImage(pictureDisp As stdole.IPictureDisp) As System.Drawing.Image
	Dim oImage As System.Drawing.Image
	If pictureDisp IsNot Nothing Then
		If pictureDisp.Type = 1 Then
			Dim hpalette As IntPtr = New IntPtr(pictureDisp.hPal)
			oImage = oImage.FromHbitmap(New IntPtr(pictureDisp.Handle), hpalette)
		End If
		If pictureDisp.Type = 2 Then
			oImage = New System.Drawing.Imaging.Metafile(New IntPtr(pictureDisp.Handle), New System.Drawing.Imaging.WmfPlaceableFileHeader())
		End If
	End If
	Return oImage
End Function
```

```vb
Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Imaging
Imports System.Windows.Graphics
Imports System.Drawing
Imports Microsoft.VisualBasic
AddReference "System.Drawing.dll"
'AddReference "Microsoft.VisualBasic.Compatibility.dll"
AddReference "Microsoft.VisualBasic.dll"
AddReference "stdole.dll"


'--------------------------------------------------------------------------------------

Sub Main()
	
	'--------------------------------------------------------------------------------------
	
	Dim oThumbnail As IPictureDisp = ThisDoc.Document.Thumbnail
	
	'--------------------------------------------------------------------------------------
	
	' oForm
	Dim oForm As New System.Windows.Forms.Form
	oForm.Text = "Thumbnail Inventor 2026"
	oForm.Size = New Drawing.Size(350, 400)
	oForm.Font = New Drawing.Font("Arial", 10)
	' oForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
	oForm.StartPosition = FormStartPosition.CenterScreen
	oForm.ShowIcon = False
	oForm.HelpButton = False
	oForm.MaximizeBox = True
	oForm.MinimizeBox = True
	oForm.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vonkajšie: Left, Top, Right, Bottom
	oForm.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	
	'--------------------------------------------------------------------------------------
	
	' TableLayoutPanel1
	Dim TableLayoutPanel1 As New System.Windows.Forms.TableLayoutPanel
	TableLayoutPanel1.ColumnCount = 1
	TableLayoutPanel1.RowCount = 1
	TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
	TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
	TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	TableLayoutPanel1.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vnútorné: Left, Top, Right, Bottom
	oForm.Controls.Add(TableLayoutPanel1)
	
	'--------------------------------------------------------------------------------------
	
	' oPictureBox
	Dim oPictureBox As New System.Windows.Forms.PictureBox
	oPictureBox.SizeMode = PictureBoxSizeMode.Zoom ' CenterImage
	oPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
	oPictureBox.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5) ' Vonkajšie: Left, Top, Right, Bottom
	oPictureBox.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0) ' Vnútorné: Left, Top, Right, Bottom
	TableLayoutPanel1.Controls.Add(oPictureBox, 0, 0)
	
	Dim oThumbnail_Image As System.Drawing.Image
	oThumbnail_Image = PictureDispToImage(oThumbnail)
	Dim oThumbnail_Bitmap As New Bitmap(500, 500)
	Dim oThumbnail_Resize As System.Drawing.Graphics = Graphics.FromImage(oThumbnail_Bitmap)
	
	Try
		
		oThumbnail_Resize.DrawImage(oThumbnail_Image, 0, 0, 500, 500)
		
	Catch ex As Exception
		
		'MsgBox(ex.ToString)
		
	End Try
	
	oPictureBox.Image = oThumbnail_Bitmap
	oThumbnail_Image = Nothing
	oThumbnail_Resize = Nothing
	oThumbnail_Bitmap = Nothing
	
	'--------------------------------------------------------------------------------------
	
	oForm.ShowDialog()
	
	'--------------------------------------------------------------------------------------
	
End Sub

Public Shared Function PictureDispToImage(pictureDisp As stdole.IPictureDisp) As System.Drawing.Image
	Dim oImage As System.Drawing.Image
	If pictureDisp IsNot Nothing Then
		If pictureDisp.Type = 1 Then
			Dim hpalette As IntPtr = New IntPtr(pictureDisp.hPal)
			oImage = oImage.FromHbitmap(New IntPtr(pictureDisp.Handle), hpalette)
		End If
		If pictureDisp.Type = 2 Then
			oImage = New System.Drawing.Imaging.Metafile(New IntPtr(pictureDisp.Handle), New System.Drawing.Imaging.WmfPlaceableFileHeader())
		End If
	End If
	Return oImage
End Function
```

---

## Invisible / abandoned weld seam symbols in drawing

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/invisible-abandoned-weld-seam-symbols-in-drawing/td-p/13837112](https://forums.autodesk.com/t5/inventor-programming-forum/invisible-abandoned-weld-seam-symbols-in-drawing/td-p/13837112)

**Author:** ThoStt

**Date:** ‎10-04-2025
	
		
		07:59 AM

**Description:** Hello, I am trying to cycle with an ilogic script through all weld symbols in a drawing. This script should show the position of each weld symbol in a messagebox.So far, the script works fine. However, it also displays the position or other properties of weld symbols that are not visible on the drawing. These weld symbols are also not located on hidden layers. I have the feeling that these weld symbols were previously assigned to deleted geometries and are still stored somewhere in the IDW.Is th...

**Code:**

```vb
Dim oActiveDoc As DrawingDocument = ThisApplication.ActiveDocument

For Each oSheet As Sheet In oActiveDoc.Sheets

	For Each oDrawingWeldsymbol As DrawingWeldingSymbol In oSheet.WeldingSymbols

			MsgBox(oDrawingWeldsymbol.Position.X & vbCrLf & oDrawingWeldsymbol.Position.Y)
			
	Next

Next
```

```vb
Dim oActiveDoc As DrawingDocument = ThisApplication.ActiveDocument

For Each oSheet As Sheet In oActiveDoc.Sheets

	For Each oDrawingWeldsymbol As DrawingWeldingSymbol In oSheet.WeldingSymbols

			MsgBox(oDrawingWeldsymbol.Position.X & vbCrLf & oDrawingWeldsymbol.Position.Y)
			
	Next

Next
```

---

## Workflow to modify MaterialAsset (apply another PhysicalPropertiesAsset)

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/workflow-to-modify-materialasset-apply-another/td-p/13822777](https://forums.autodesk.com/t5/inventor-programming-forum/workflow-to-modify-materialasset-apply-another/td-p/13822777)

**Author:** Maxim-CADman77

**Date:** ‎09-23-2025
	
		
		04:12 PM

**Description:** I'm studying the possibility to automate modification of MaterialAssets (apply another PhysicalPropertiesAsset according to some logic).PhysicalPropertiesAsset description says it is a writable Asset property and "When assigning physical properties, the physical properties asset must exist in the same document as the material." Ok. I do have two materials target and source in my Part document.I then try to assign PhysicalPropertiesAsset of source material to PhysicalPropertiesAsset of target mat...

**Code:**

```vb
Dim ptDoc As PartDocument = ThisDoc.Document

Logger.Info(ptDoc.MaterialAssets.Count)

Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Logger.Info(srcMat.DisplayName)
Logger.Info(vbTab & srcMat.PhysicalPropertiesAsset.DisplayName)

Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
Logger.Info(tgtMat.DisplayName)
Logger.Info(vbTab & tgtMat.PhysicalPropertiesAsset.DisplayName)

tgtMat.PhysicalPropertiesAsset = srcMat.PhysicalPropertiesAsset ' HRESULT: 0x80020003 (DISP_E_MEMBERNOTFOUND)
```

```vb
Dim ptDoc As PartDocument = ThisDoc.Document

Logger.Info(ptDoc.MaterialAssets.Count)

Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Logger.Info(srcMat.DisplayName)
Logger.Info(vbTab & srcMat.PhysicalPropertiesAsset.DisplayName)

Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
Logger.Info(tgtMat.DisplayName)
Logger.Info(vbTab & tgtMat.PhysicalPropertiesAsset.DisplayName)

tgtMat.PhysicalPropertiesAsset = srcMat.PhysicalPropertiesAsset ' HRESULT: 0x80020003 (DISP_E_MEMBERNOTFOUND)
```

```vb
Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
```

```vb
Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
```

```vb
Dim srcMat As MaterialAsset = ptDoc.MaterialAssets(1)
Dim tgtMat As MaterialAsset = ptDoc.MaterialAssets(2)
```

```vb
Dim srcMat As MaterialAsset = ptDoc.MaterialAssets(1)
Dim tgtMat As MaterialAsset = ptDoc.MaterialAssets(2)
```

---

## Error 0x800AC472

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/error-0x800ac472/td-p/13755364](https://forums.autodesk.com/t5/inventor-programming-forum/error-0x800ac472/td-p/13755364)

**Author:** layochim

**Date:** ‎08-05-2025
	
		
		05:59 PM

**Description:** I'm still struggling with this iLogic error when writing to an Excel file and pulling a Value from that Excel to a parameter in my model.  Are there any known solutions to 0x800AC472?  My code is below, I've tried using the full file path name and with no success in removing the error.  The rule will successfully pull a value into my parameter the first time then error out after the second. Dim Filename As String = "Straight Conveyor Pricing.xlsx"
Dim SheetName As String = "Conv costing"

GoExce...

**Code:**

```vb
Dim Filename As String = "Straight Conveyor Pricing.xlsx"
Dim SheetName As String = "Conv costing"

GoExcel.Open(Filename, SheetName)

GoExcel.CellValue("B9") = BeltWidth
GoExcel.CellValue("B10") = BeltLength
GoExcel.CellValue("B11") = Infeed_End
GoExcel.CellValue("B12") = Discharge_End
GoExcel.CellValue("B13") = Legs
GoExcel.CellValue("B14") = Feet
GoExcel.CellValue("B15") = Belting
GoExcel.CellValue("B16") = BeltingType
GoExcel.CellValue("B17") = BeltColor
GoExcel.CellValue("B18") = Flights
GoExcel.CellValue("B19") = Flight_Spacing
GoExcel.CellValue("B20") = Construction
GoExcel.CellValue("B21") = FlangeHeight
GoExcel.CellValue("B22") = Finish
GoExcel.CellValue("B24") = Drive_Style
GoExcel.CellValue("B25") = Drive_Position
GoExcel.CellValue("B26") = MotorVoltage
GoExcel.CellValue("B27") = Drive_Speed
GoExcel.CellValue("B28") = Motor
GoExcel.CellValue("B29") = Gearbox
GoExcel.CellValue("B30") = AdjustableGuides
GoExcel.CellValue("B31") = BeltLifters
GoExcel.CellValue("B32") = DripPan
GoExcel.CellValue("B33") = BeltScraper
GoExcel.CellValue("B34") = SprayBar
GoExcel.CellValue("B35") = GearboxGuards

Price = GoExcel.CellValue("B6")
GoExcel.Save
GoExcel.Close
```

```vb
Dim oExcelPath As Object = "C:\Users\Lyle\SCR Solutions\Shared OneDrive - Draftsmen Access\Configurator files\Straight Conveyor Configurator\CAD\Straight Conveyor Pricing Test.xlsx"
'	Dim oExcelPath As String = "Straight Conveyor Pricing.xlsx"
	Dim oExcel As Object = CreateObject("Excel.Application")
	oExcel.Visible = False
	oExcel.DisplayAlerts = False
	Dim oWB As Object = oExcel.Workbooks.Open(oExcelPath)
	Dim oWS As Object = oWB.Sheets(1)
	oWS.Activate


oWS.Cells(9, 2) = BeltWidth
oWS.Cells(10, 2).Value = BeltLength
oWS.Cells(11, 2).Value = Infeed_End
oWS.Cells(12, 2).Value = Discharge_End
oWS.Cells(13, 2).Value = Legs
oWS.Cells(14, 2).Value = Feet
oWS.Cells(15, 2).Value = Belting
'oWS.Cells(16, 2).Value = BeltingType
'oWS.Cells(17, 2).Value = BeltColor
'oWS.Cells(18, 2).Value = Flights
'oWS.Cells(19, 2).Value = Flight_Spacing
oWS.Cells(20, 2).Value = Construction
'oWS.Cells(21, 2).Value = FlangeHeight
oWS.Cells(22, 2).Value = Finish
oWS.Cells(24, 2).Value = Drive_Style
'oWS.Cells(25, 2).Value = Drive_Position
'oWS.Cells(26, 2).Value = MotorVoltage
'oWS.Cells(27, 2).Value = Drive_Speed
oWS.Cells(28, 2).Value = Motor
oWS.Cells(29, 2).Value = Gearbox
oWS.Cells(30, 2).Value = AdjustableGuides
oWS.Cells(31, 2).Value = BeltLifters
oWS.Cells(32, 2).Value = DripPan
'oWS.Cells(33, 2).Value = BeltScraper
oWS.Cells(34, 2).Value = SprayBar
oWS.Cells(35, 2).Value = GearboxGuards

Price = oWS.Cells(6, 2).Value
	oWB.Close (True)
	oExcel.Quit
	oExcel = Nothing
```

```vb
Dim Filename As String = "Straight Conveyor Pricing Test2.xlsx"
Dim SheetName As String = "Conv costing"
```

---

## Ilogic split a string

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436#messageview_0)

**Author:** Anonymous

**Date:** ‎05-08-2019
	
		
		12:20 AM

**Description:** Goodday,
I am searching for a way to split strings. Where:
Dim Separators() As Char = {";"} 
Doesnt seem to work at my code.
 
As you see in the code the separator is-> ;
for example, if the string is 1;2, it must be separated into 1 and 2 given to 2 parameters,
so spot1 -> "1" and spot2 -> "2"
This also applies to, for example, 4;6;13 in 4, 6 and 13 given to 3 parameters, etc.
 
I hope i can find the solution im looking for over here!
Kind Regards,
Thomas
					
				
			
			
				
			
			
				
	
...

**Code:**

```vb
Dim Separators() As Char = {";"}
```

```vb
Dim Separators() As Char = {";"} 
Sentence = "1;2;3;4;5"
Words = Sentence.Split(Separators)
i = 0
For Each wrd In Words
MessageBox.Show("Word Index #" & i & " = " & Words(i))
i=i+1
Next
```

```vb
Dim Separators() As Char = {";"} 
Sentence = "10;20;3;4;5"
Words = Sentence.Split(Separators)

spot1 = Words(0)
spot2 = Words(1)

Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
```

```vb
Dim Separators() As Char = {";"} 
Sentence = StringParameter
If Sentence = "1;1" Then
Words = Sentence.Split(Separators)
spot1 = Words(0)
spot2 = Words(1)
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
Else If Sentence = "1;5;7" Then
Words = Sentence.Split(Separators)
spot1 = Words(0)
spot2 = Words(1)
spot3 = Words(2)
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
End If 
InventorVb.DocumentUpdate()
```

```vb
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
```

```vb
Dim Separators() As Char = {";",vbCrLf} Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(Separators)
i = 0
For Each wrd In Words
MessageBox.Show("Word Index #" & i & " = " & Words(i))
i=i+1
Next
```

```vb
Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(vbCrLf)

For Each wrd In Words
	'strip out carraige returns & linefeeds
	wrd = wrd.Replace(vbCr, "").Replace(vbLf, "")
	MessageBox.Show("Word Index #" & i & " = " & wrd)
Next
```

```vb
Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(vbCrLf)

For Each wrd In Words
	'strip out carraige returns & linefeeds
	wrd = wrd.Replace(vbCr, "").Replace(vbLf, "")
	MessageBox.Show("Word Index #" & i & " = " & wrd)
Next
```

```vb
Sentence = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
```

```vb
Sentence = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
```

```vb
Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Separators() As Char = {".", "-", " " }
Dim Words() As String = Sentence.Split(Separators)
Dim Word1 As String = Words(0)
Dim Word2 As String = Words(1)
Dim Word3 As String = Words(2)
Dim Word4 As String = Words(3)
Dim Word5 As String = Words(4)
Dim Word6 As String = Words(5)
MsgBox("Word1 = " & Word1 & _
vbCrLf & "Word2 = " & Word2 & _
vbCrLf & "Word3 = " & Word3 & _
vbCrLf & "Word4 = " & Word4 & _
vbCrLf & "Word5 = " & Word5 & _
vbCrLf & "Word6 = " & Word6, vbInformation, "Sentence Split Into Words")
```

```vb
Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Words() As String = Sentence.Split({".", "-", " " },StringSplitOptions.RemoveEmptyEntries)
Dim sResults As String
For i As Integer = 0 To Words.Length - 1
	sResults &= "Word " & i & " = " & Words(i) & vbCrLf
Next
MsgBox(sResults, vbInformation, "Sentence Split Into Words")
```

```vb
Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Words() As String = Sentence.Split({".", "-", " " },StringSplitOptions.RemoveEmptyEntries)
Dim sResults As String
For i As Integer = 0 To Words.Length - 1
	sResults &= "Word " & i & " = " & Words(i) & vbCrLf
Next
MsgBox(sResults, vbInformation, "Sentence Split Into Words")
```

---

## How to &quot;Show All&quot; in a part document with iLogic?

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-show-all-quot-in-a-part-document-with-ilogic/td-p/13832701#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-show-all-quot-in-a-part-document-with-ilogic/td-p/13832701#messageview_0)

**Author:** acheungR3ZK5

**Date:** ‎10-01-2025
	
		
		04:46 AM

**Description:** How do I call the "Show All" command (shown in screenshot) in a part document with iLogic? I know how to loop through all the parts and make them visible if not visible, but it takes longer if there are many parts and you can't undo the whole command in one undo.  
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim partDoc As PartDocument
partDoc = ThisApplication.ActiveDocument

Dim compDef As PartComponentDefinition
compDef = partDoc.ComponentDefinition

Dim activeViewRep As DesignViewRepresentation = compDef.RepresentationsManager.ActiveDesignViewRepresentation
activeViewRep.ShowAll
```

```vb
Dim partDoc As PartDocument
partDoc = ThisApplication.ActiveDocument

Dim compDef As PartComponentDefinition
compDef = partDoc.ComponentDefinition

Dim activeViewRep As DesignViewRepresentation = compDef.RepresentationsManager.ActiveDesignViewRepresentation
activeViewRep.ShowAll
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("ShowAllBodiesCtxCmd").Execute()
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("ShowAllBodiesCtxCmd").Execute()
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("HideOtherBodiesCtxCmd").Execute()
```

```vb
ThisApplication.CommandManager.ControlDefinitions.Item("HideOtherBodiesCtxCmd").Execute()
```

```vb
Dim oApp As Inventor.Application = ThisApplication
Dim oPDoc As PartDocument = oApp.ActiveDocument
Dim oSBColl As ObjectCollection = oApp.TransientObjects.CreateObjectCollection()
For Each oSB As SurfaceBody In oPDoc.ComponentDefinition.SurfaceBodies
    oSBColl.Add(oSB)
Next
oPDoc.SelectSet.Clear()
oPDoc.SelectSet.SelectMultiple(oSBColl)
oApp.CommandManager.ControlDefinitions.Item("PartVisibilityCtxCmd").Execute()
```

```vb
Dim oApp As Inventor.Application = ThisApplication
Dim oPDoc As PartDocument = oApp.ActiveDocument
Dim oSBColl As ObjectCollection = oApp.TransientObjects.CreateObjectCollection()
For Each oSB As SurfaceBody In oPDoc.ComponentDefinition.SurfaceBodies
    oSBColl.Add(oSB)
Next
oPDoc.SelectSet.Clear()
oPDoc.SelectSet.SelectMultiple(oSBColl)
oApp.CommandManager.ControlDefinitions.Item("PartVisibilityCtxCmd").Execute()
```

---

## Macro to delete specific sketch symbol?

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606#messageview_0)

**Author:** jacob_r_kane

**Date:** ‎08-26-2025
	
		
		08:37 AM

**Description:** Hello,Without spending days trying to learn this myself, what is the code to scan all drawing sheets, and delete a sketch symbol named rev triangle?Thanks
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, Inventor.DrawingDocument)
If oDDoc Is Nothing Then Return
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oSS As SketchedSymbol In oSheet.SketchedSymbols
		If oSS.Name = "rev triangle" Then
			oSS.Delete()
		End If
	Next 'oSS
Next 'oSheet
```

```vb
Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, Inventor.DrawingDocument)
If oDDoc Is Nothing Then Return
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oSS As SketchedSymbol In oSheet.SketchedSymbols
		If oSS.Name = "rev triangle" Then
			oSS.Delete()
		End If
	Next 'oSS
Next 'oSheet
```

```vb
Sub DeleteRevTriangleSketchedSymbols()
    If Not ThisApplication.ActiveDocumentType = kDrawingDocumentObject Then Return
    Dim oDDoc As DrawingDocument
    Set oDDoc = ThisApplication.ActiveDocument
    Dim oSheet As Inventor.Sheet
    For Each oSheet In oDDoc.Sheets
        Dim oSS As SketchedSymbol
        For Each oSS In oSheet.SketchedSymbols
            If oSS.Name = "rev triangle" Then
                Call oSS.Delete
            End If
        Next 'oSS
    Next 'oSheet
End Sub
```

```vb
Sub DeleteRevTriangleSketchedSymbols()
    If Not ThisApplication.ActiveDocumentType = kDrawingDocumentObject Then Return
    Dim oDDoc As DrawingDocument
    Set oDDoc = ThisApplication.ActiveDocument
    Dim oSheet As Inventor.Sheet
    For Each oSheet In oDDoc.Sheets
        Dim oSS As SketchedSymbol
        For Each oSS In oSheet.SketchedSymbols
            If oSS.Name = "rev triangle" Then
                Call oSS.Delete
            End If
        Next 'oSS
    Next 'oSheet
End Sub
```

```vb
Public Sub DeleteSpecificSketchSymbol()
    ' Set a reference to the active drawing document.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    ' Define the name of the sketch symbol to delete.
    Const sSymbolName As String = "Rev Triangle" ' Replace with the actual symbol name
    ' Loop through each sheet in the drawing document.
    Dim oSheet As Sheet
    For Each oSheet In oDrawDoc.Sheets
        ' Activate the current sheet (optional, but good practice for visibility).
        oSheet.Activate
        ' Loop through each sketched symbol on the current sheet.
        Dim oSketchedSymbol As SketchedSymbol
        Dim i As Long
        For i = oSheet.SketchedSymbols.Count To 1 Step -1 ' Loop backwards to avoid issues when deleting
            Set oSketchedSymbol = oSheet.SketchedSymbols.Item(i)
            ' Check if the symbol's name matches the target name.
            If oSketchedSymbol.Name = sSymbolName Then
                ' Delete the instance of the sketch symbol.
                oSketchedSymbol.Delete
            End If
        Next i
    Next oSheet
End Sub
```

```vb
Public Sub DeleteSpecificSketchSymbol()
    ' Set a reference to the active drawing document.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    ' Define the name of the sketch symbol to delete.
    Const sSymbolName As String = "Rev Triangle" ' Replace with the actual symbol name
    ' Loop through each sheet in the drawing document.
    Dim oSheet As Sheet
    For Each oSheet In oDrawDoc.Sheets
        ' Activate the current sheet (optional, but good practice for visibility).
        oSheet.Activate
        ' Loop through each sketched symbol on the current sheet.
        Dim oSketchedSymbol As SketchedSymbol
        Dim i As Long
        For i = oSheet.SketchedSymbols.Count To 1 Step -1 ' Loop backwards to avoid issues when deleting
            Set oSketchedSymbol = oSheet.SketchedSymbols.Item(i)
            ' Check if the symbol's name matches the target name.
            If oSketchedSymbol.Name = sSymbolName Then
                ' Delete the instance of the sketch symbol.
                oSketchedSymbol.Delete
            End If
        Next i
    Next oSheet
End Sub
```

---

## Use filename of 2D file, to fill in iproperties - Inventor 2024

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/use-filename-of-2d-file-to-fill-in-iproperties-inventor-2024/td-p/13802821#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/use-filename-of-2d-file-to-fill-in-iproperties-inventor-2024/td-p/13802821#messageview_0)

**Author:** jens_herrebaut83

**Date:** ‎09-09-2025
	
		
		04:46 AM

**Description:** Hi,
 
Is there a way to save a 2D file (ex: Title-desciption-customer.idw) and it automatically fills in the iproperties of that 2D file?
 
 
@jens_herrebaut83 Your post title was modified to add the product name and version and to increase findability - CGBenner
 
 
					
				
			
			
				
			
			
				
	
			
				
					
						Solved!
					
					Go to Solution.

**Code:**

```vb
Try
	Dim oDoc As Document = ThisApplication.ActiveEditDocument
	Dim oName As String = ThisDoc.FileName(False)
	Dim oFirstDashPos As Integer = InStr(oName, "-")
	Dim oSecondDashPos As Integer = InStr(oFirstDashPos + 1, oName, "-")
	Dim oTitle As String = Left(oName, oFirstDashPos -1)
	Dim oClient As String = Mid(oName, oFirstDashPos + 1, oSecondDashPos - oFirstDashPos - 1)
	Dim oDesc As String = Right(oName, Len(oName) - oSecondDashPos)
	
	iProperties.Value("Summary", "Title") = oTitle
	iProperties.Value("Custom", "Client") = oClient
	iProperties.Value("Project", "Description") = oDesc
Catch
End Try
```

```vb
Try
	Dim oDoc As Document = ThisApplication.ActiveEditDocument
	Dim oName As String = ThisDoc.FileName(False)
	Dim oFirstDashPos As Integer = InStr(oName, "-")
	Dim oSecondDashPos As Integer = InStr(oFirstDashPos + 1, oName, "-")
	Dim oTitle As String = Left(oName, oFirstDashPos -1)
	Dim oClient As String = Mid(oName, oFirstDashPos + 1, oSecondDashPos - oFirstDashPos - 1)
	Dim oDesc As String = Right(oName, Len(oName) - oSecondDashPos)
	
	iProperties.Value("Summary", "Title") = oTitle
	iProperties.Value("Custom", "Client") = oClient
	iProperties.Value("Project", "Description") = oDesc
Catch
End Try
```

---

## acad.exe link problem!

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/acad-exe-link-problem/td-p/13822188](https://forums.autodesk.com/t5/inventor-programming-forum/acad-exe-link-problem/td-p/13822188)

**Author:** SergeLachance

**Date:** ‎09-23-2025
	
		
		07:43 AM

**Description:** i have a old rule who working great until a install inventor 2026??? i just change the new link but error message!any body can help me???sorry for my horrible english! my old rule: ' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveDocument
Dim oContext As T...

**Code:**

```vb
' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveDocument
Dim oContext As TranslationContext
oContext = ThisApplication.TransientObjects.CreateTranslationContext
oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
' Create a NameValueMap object
Dim oOptions As NameValueMap
oOptions = ThisApplication.TransientObjects.CreateNameValueMap
' Create a DataMedium object
Dim oDataMedium As DataMedium
oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
' Check whether the translator has 'SaveCopyAs' options
If DXFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DXFOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

DOSSIER = ThisDoc.Path
POSITION1 = InStrRev(DOSSIER, "\") 
POSITION2 = Right(DOSSIER, Len(DOSSIER) - POSITION1)

'get DXF target folder path
oFolder = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
    System.IO.Directory.CreateDirectory(oFolder)
End If

For i = 1 To 100
chiffre = i
myFile = oFolder & oFileName & "_Sheet_" & chiffre & ".dxf"
If(System.IO.File.Exists(myFile)) Then
Kill (myFile)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder & oFileName & ".dxf"
oFileName = ThisDoc.FileName(False) 'without extension

'Publish document.
DXFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder & oFileName & "_Sheet_" & 100 & ".dxf"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

' Get the DWG translator Add-In.
Dim oDWGAddIn As TranslatorAddIn 
oDWGAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
' Check whether the translator has 'SaveCopyAs' options
If oDWGAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DWGOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

'get DXF target folder path
oFolder2 = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder2) Then
    System.IO.Directory.CreateDirectory(oFolder2)
End If

For i = 1 To 100
chiffre2 = i
myFile2 = oFolder2 & oFileName2 & "_Sheet_" & chiffre2 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder2 & oFileName2 & ".dwg"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Publish document.
oDWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder2 & oFileName2 & "_Sheet_" & 100 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

question = MessageBox.Show("VEUX-TU OUVRIR AUTOCAD POUR PURGER TES DESSINS???", "OUVERTURE AUTOCAD???", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
MultiValue.SetValueOptions(True)
If question = vbYes Then
GoTo OUVERTURE
Else If question = vbNo Then
GoTo FIN
End If

OUVERTURE :
Dim acadExe = "C:\Program Files\Autodesk\AutoCAD 2025\acad.exe"

'MessageBox.Show("N'OUBLIE PAS DE SAUVER EN VERSION 2000, UNE FOIS TON DESSIN PURGER ET DE PURGER TOUTES TES PAGES", "ATTENTION",MessageBoxButtons.OK,MessageBoxIcon.Warning)
Process.Start(acadExe, oFolder2 & oFileName2 & "_Sheet_" & 1 & ".dwg")

FIN:
```

```vb
' Get the DXF translator Add-In.
Dim DXFAddIn As TranslatorAddIn
DXFAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
'Set a reference to the active document (the document to be published).
Dim oDocument As Document
oDocument = ThisApplication.ActiveDocument
Dim oContext As TranslationContext
oContext = ThisApplication.TransientObjects.CreateTranslationContext
oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
' Create a NameValueMap object
Dim oOptions As NameValueMap
oOptions = ThisApplication.TransientObjects.CreateNameValueMap
' Create a DataMedium object
Dim oDataMedium As DataMedium
oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
' Check whether the translator has 'SaveCopyAs' options
If DXFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DXFOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

DOSSIER = ThisDoc.Path
POSITION1 = InStrRev(DOSSIER, "\") 
POSITION2 = Right(DOSSIER, Len(DOSSIER) - POSITION1)

'get DXF target folder path
oFolder = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
    System.IO.Directory.CreateDirectory(oFolder)
End If

For i = 1 To 100
chiffre = i
myFile = oFolder & oFileName & "_Sheet_" & chiffre & ".dxf"
If(System.IO.File.Exists(myFile)) Then
Kill (myFile)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder & oFileName & ".dxf"
oFileName = ThisDoc.FileName(False) 'without extension

'Publish document.
DXFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder & oFileName & "_Sheet_" & 100 & ".dxf"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

' Get the DWG translator Add-In.
Dim oDWGAddIn As TranslatorAddIn 
oDWGAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")
' Check whether the translator has 'SaveCopyAs' options
If oDWGAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
Dim strIniFile As String
strIniFile = "I:\INVENTOR\ILOGIC\TEMPLATE\DWGOut.ini"
' Create the name-value that specifies the ini file to use.
oOptions.Value("Export_Acad_IniFile") = strIniFile
End If

'get DXF target folder path
oFolder2 = "I:\DOCUMENTS\ARTICLES\" & POSITION2 & "\"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Check for the PDF folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder2) Then
    System.IO.Directory.CreateDirectory(oFolder2)
End If

For i = 1 To 100
chiffre2 = i
myFile2 = oFolder2 & oFileName2 & "_Sheet_" & chiffre2 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If
Next i

'Set the destination file name
oDataMedium.FileName = oFolder2 & oFileName2 & ".dwg"
oFileName2 = ThisDoc.FileName(False) 'without extension

'Publish document.
oDWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'Launch the dxf file in whatever application Windows is set to open this document type with

myFile2 = oFolder2 & oFileName2 & "_Sheet_" & 100 & ".dwg"
If(System.IO.File.Exists(myFile2)) Then
Kill (myFile2)
End If

question = MessageBox.Show("VEUX-TU OUVRIR AUTOCAD POUR PURGER TES DESSINS???", "OUVERTURE AUTOCAD???", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
MultiValue.SetValueOptions(True)
If question = vbYes Then
GoTo OUVERTURE
Else If question = vbNo Then
GoTo FIN
End If

OUVERTURE :
Dim acadExe = "‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"

'MessageBox.Show("N'OUBLIE PAS DE SAUVER EN VERSION 2000, UNE FOIS TON DESSIN PURGER ET DE PURGER TOUTES TES PAGES", "ATTENTION",MessageBoxButtons.OK,MessageBoxIcon.Warning)
Process.Start(acadExe, oFolder2 & oFileName2 & "_Sheet_" & 1 & ".dwg")

FIN:
```

```vb
Dim acadExe = "‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"

'MessageBox.Show("N'OUBLIE PAS DE SAUVER EN VERSION 2000, UNE FOIS TON DESSIN PURGER ET DE PURGER TOUTES TES PAGES", "ATTENTION",MessageBoxButtons.OK,MessageBoxIcon.Warning)
Process.Start(acadExe)
```

```vb
"‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"
```

```vb
Dim acadExe = "‪‪C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"

Process.Start(acadExe)
```

```vb
Process.Start(acad.exe)
```

```vb
C:\Program Files\Autodesk\AutoCAD 2026\acad.exe
```

```vb
Process.Start("notepad.exe")
```

---

## iLogic Drawing Layer Visibility Toggle

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-drawing-layer-visibility-toggle/td-p/13816579#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-drawing-layer-visibility-toggle/td-p/13816579#messageview_0)

**Author:** dave_novak

**Date:** ‎09-18-2025
	
		
		10:09 AM

**Description:** Hello, I would like to figure out (with a little help from my friends) how to get Drawings to show the correct Visibility State in the Layer Drop Down after running the following iLogic Rule: If ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True
	ElseIf ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True Then
	ThisDrawing.Document.StylesManager.Lay...

**Code:**

```vb
If ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True
	ElseIf ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False
	
End If

iLogicVb.UpdateWhenDone = True
```

---

## Have the sheet name automatically generate in Title block

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336](https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336)

**Author:** BrandonLee_Innovex

**Date:** ‎11-26-2021
	
		
		12:06 PM

**Description:** I am trying to have my title block sheet name automatically populated once the sheet name has been set.  For example, the image i have attached shows the sheet names we type in, but our current process requires us to type in the sheet name multiple times to appear both in the sidebar shown as well as the title block. I am wondering if there is a way to make the title block auto-populate once the sheet name has been filled out in the sidebar. 
					
				
			
			
				
			
			
				
	
			
				
					...

**Code:**

```vb
Dim doc As DrawingDocument = ThisDoc.Document
Dim sheet As Sheet = doc.ActiveSheet
Dim titleBlock As TitleBlock = sheet.TitleBlock

Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
        Cast(Of Inventor.TextBox).
        Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
        Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
        FirstOrDefault()
If (textBox Is Nothing) Then Throw New Exception("Promted entry not found.")
titleBlock.SetPromptResultText(textBox, sheet.Name)
```

```vb
Dim doc As DrawingDocument = ThisDoc.Document
Dim sheet As Sheet = doc.ActiveSheet
Dim titleBlock As TitleBlock = sheet.TitleBlock

Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
        Cast(Of Inventor.TextBox).
        Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
        Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
        FirstOrDefault()
If (textBox Is Nothing) Then Throw New Exception("Promted entry not found.")

Dim sheetDescription = sheet.Name
Dim puntPlace = InStr(sheetDescription, ":") - 1
sheetDescription = sheetDescription.Substring(0, puntPlace)
titleBlock.SetPromptResultText(textBox, sheetDescription)
```

```vb
Dim doc As DrawingDocument = ThisDoc.Document
For Each sheet As Sheet In doc.Sheets
    Dim titleBlock As TitleBlock = Sheet.TitleBlock

    Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
	    Cast(Of Inventor.TextBox).
	    Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
	    Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
	    FirstOrDefault()
    If (textBox Is Nothing) Then 
		MsgBox("Promted entry not found on sheet: " & Sheet.Name)
		continue for
	End If

    Dim sheetDescription = Sheet.Name
    Dim puntPlace = InStr(sheetDescription, ":") - 1
    sheetDescription = sheetDescription.Substring(0, puntPlace)
    titleBlock.SetPromptResultText(textBox, sheetDescription)
Next
```

```vb
Dim doc As DrawingDocument = ThisDoc.Document
For Each sheet As Sheet In doc.Sheets
    Dim titleBlock As TitleBlock = Sheet.TitleBlock

    ' Find the textbox within the title block's sketch that contains the prompt markers
    Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes _
        .Cast(Of Inventor.TextBox)() _
        .Where(Function(tb) tb.FormattedText.Contains("<Prompt") AndAlso tb.FormattedText.Contains("SHEET TITLE")) _
        .FirstOrDefault()

    If (textBox Is Nothing) Then
        MsgBox("Prompted entry not found on sheet: " & Sheet.Name)
        Continue For
    End If

    ' Extract the sheet description; here, trimming after colon
    Dim sheetDescription As String = Sheet.Name
    Dim puntPlace As Integer = InStr(sheetDescription, ":")
    If puntPlace > 0 Then
        sheetDescription = sheetDescription.Substring(0, puntPlace - 1)
    End If

    ' Set the prompt result text in the title block
    titleBlock.SetPromptResultText(textBox, sheetDescription)
Next
```

---

## Have the sheet name automatically generate in Title block

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336#messageview_0)

**Author:** BrandonLee_Innovex

**Date:** ‎11-26-2021
	
		
		12:06 PM

**Description:** I am trying to have my title block sheet name automatically populated once the sheet name has been set.  For example, the image i have attached shows the sheet names we type in, but our current process requires us to type in the sheet name multiple times to appear both in the sidebar shown as well as the title block. I am wondering if there is a way to make the title block auto-populate once the sheet name has been filled out in the sidebar. 
					
				
			
			
				
			
			
				
	
			
				
					...

**Code:**

```vb
Dim doc As DrawingDocument = ThisDoc.Document
Dim sheet As Sheet = doc.ActiveSheet
Dim titleBlock As TitleBlock = sheet.TitleBlock

Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
        Cast(Of Inventor.TextBox).
        Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
        Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
        FirstOrDefault()
If (textBox Is Nothing) Then Throw New Exception("Promted entry not found.")
titleBlock.SetPromptResultText(textBox, sheet.Name)
```

```vb
Dim doc As DrawingDocument = ThisDoc.Document
Dim sheet As Sheet = doc.ActiveSheet
Dim titleBlock As TitleBlock = sheet.TitleBlock

Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
        Cast(Of Inventor.TextBox).
        Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
        Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
        FirstOrDefault()
If (textBox Is Nothing) Then Throw New Exception("Promted entry not found.")

Dim sheetDescription = sheet.Name
Dim puntPlace = InStr(sheetDescription, ":") - 1
sheetDescription = sheetDescription.Substring(0, puntPlace)
titleBlock.SetPromptResultText(textBox, sheetDescription)
```

```vb
Dim doc As DrawingDocument = ThisDoc.Document
For Each sheet As Sheet In doc.Sheets
    Dim titleBlock As TitleBlock = Sheet.TitleBlock

    Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
	    Cast(Of Inventor.TextBox).
	    Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
	    Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
	    FirstOrDefault()
    If (textBox Is Nothing) Then 
		MsgBox("Promted entry not found on sheet: " & Sheet.Name)
		continue for
	End If

    Dim sheetDescription = Sheet.Name
    Dim puntPlace = InStr(sheetDescription, ":") - 1
    sheetDescription = sheetDescription.Substring(0, puntPlace)
    titleBlock.SetPromptResultText(textBox, sheetDescription)
Next
```

```vb
Dim doc As DrawingDocument = ThisDoc.Document
For Each sheet As Sheet In doc.Sheets
    Dim titleBlock As TitleBlock = Sheet.TitleBlock

    ' Find the textbox within the title block's sketch that contains the prompt markers
    Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes _
        .Cast(Of Inventor.TextBox)() _
        .Where(Function(tb) tb.FormattedText.Contains("<Prompt") AndAlso tb.FormattedText.Contains("SHEET TITLE")) _
        .FirstOrDefault()

    If (textBox Is Nothing) Then
        MsgBox("Prompted entry not found on sheet: " & Sheet.Name)
        Continue For
    End If

    ' Extract the sheet description; here, trimming after colon
    Dim sheetDescription As String = Sheet.Name
    Dim puntPlace As Integer = InStr(sheetDescription, ":")
    If puntPlace > 0 Then
        sheetDescription = sheetDescription.Substring(0, puntPlace - 1)
    End If

    ' Set the prompt result text in the title block
    titleBlock.SetPromptResultText(textBox, sheetDescription)
Next
```

---

## VBA : Using GetObject in Inventor 2024

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/vba-using-getobject-in-inventor-2024/td-p/13777048](https://forums.autodesk.com/t5/inventor-programming-forum/vba-using-getobject-in-inventor-2024/td-p/13777048)

**Author:** TONELLAL

**Date:** ‎08-21-2025
	
		
		04:05 AM

**Description:** Hello,I have several macros in VBA, using Excel. They connect to Excel using GetObject(, "Excel.application"), or CreateObject("Excel.application").These commands are not existing anymore in .NET framework > 5, so they don't work anymore with Inventor 2024.I found several links using .NET or C#, but noting on VBA.How can I modify these macros, still in VBA, so that they work with version 2024 and future versions of Inventor?

**Code:**

```vb
Option Explicit
Sub OpenExcel()
    Dim oExcelApp As Object
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible = True
    
    Dim oWB As Object
    Set oWB = oExcelApp.Workbooks.Add()
    
    Dim oWS As Object
    Set oWS = oWB.Worksheets.Add()
    oWS.Name = "My New Worksheet"
    
    Call MsgBox("Review New Instance Of Excel, And New, Renamed Sheet In It.", , "")
    
    Call oWB.Close
    Call oExcelApp.Quit
    
    Set oWS = Nothing
    Set oWB = Nothing
    Set oExcelApp = Nothing
End Sub
```

```vb
Option Explicit
Sub OpenExcel()
    Dim oExcelApp As Object
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible = True
    
    Dim oWB As Object
    Set oWB = oExcelApp.Workbooks.Add()
    
    Dim oWS As Object
    Set oWS = oWB.Worksheets.Add()
    oWS.Name = "My New Worksheet"
    
    Call MsgBox("Review New Instance Of Excel, And New, Renamed Sheet In It.", , "")
    
    Call oWB.Close
    Call oExcelApp.Quit
    
    Set oWS = Nothing
    Set oWB = Nothing
    Set oExcelApp = Nothing
End Sub
```

---

## Copy Design Issue – Suppressed Parts Not Rebuilding or Updating (Skeleton Reference Problem)

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/copy-design-issue-suppressed-parts-not-rebuilding-or-updating/td-p/13816281](https://forums.autodesk.com/t5/inventor-programming-forum/copy-design-issue-suppressed-parts-not-rebuilding-or-updating/td-p/13816281)

**Author:** pranaykapse1

**Date:** ‎09-18-2025
	
		
		06:46 AM

**Description:** Hi Everyone, I am facing an issue while working with Copy Design in Vault for my Autodesk Inventor assemblies.All of my assemblies are parametric and skeleton-based. The Copy Design process works fine for normal parts and assemblies, but I encounter problems with suppressed components.When suppressed parts are included in the Copy Design, they often retain references to the original skeleton file instead of re-linking to the newly copied skeleton. To fix this, I currently need to:Open each suppr...

**Code:**

```vb
'This Rule Helps to Update all parts and assemblies (Even Suppressed) and Save.

For Each oDoc As Document In ThisApplication.Documents
Try
' Force update
oDoc.Update()

' Save after update
If oDoc.FullFileName <> "" AndAlso oDoc.IsModifiable Then
oDoc.Save2(True) ' True = Save silently without dialogs
End If
Catch
End Try
Next
iLogicVb.UpdateWhenDone = True
```

```vb
'This Rule Helps to Update all parts and assemblies (Even Suppressed) and Save.

For Each oDoc As Document In ThisApplication.Documents
Try
' Force update
oDoc.Update()

' Save after update
If oDoc.FullFileName <> "" AndAlso oDoc.IsModifiable Then
oDoc.Save2(True) ' True = Save silently without dialogs
End If
Catch
End Try
Next
iLogicVb.UpdateWhenDone = True
```

---

## How to refresh/open the projects window VBA

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/how-to-refresh-open-the-projects-window-vba/td-p/13725927#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-refresh-open-the-projects-window-vba/td-p/13725927#messageview_0)

**Author:** gablewatson44

**Date:** ‎07-15-2025
	
		
		07:37 AM

**Description:** Currently using inventor 2024, i have written a macro that automatically creates a workgroup path based on the file location of whatever is opened. I am using a work around with Sendkeys that quickly opens and closes the project window using a custom shortcut:' Type P, R, O then EscapeSendKeys "P", TrueSendKeys "R", TrueSendKeys "O", TrueSendKeys "{ESC}", True while this works, i would rather use a form to completely control this macro thus requiring a different method of refreshing the project ...

**Code:**

```vb
Dim oControlDef As ControlDefinition
        
Set oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppProjectsCmd")
        
Call oControlDef.Execute
```

```vb
Dim oControlDef As ControlDefinition
        
Set oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppProjectsCmd")
        
Call oControlDef.Execute
```

---

## ControlDefinition.Execute2 does not run unless it is the last line in the rule

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361#messageview_0)

**Author:** Nick_Hall

**Date:** ‎10-05-2025
	
		
		09:06 PM

**Description:** I am writing some iLogic that interacts with the Vault Revision Table. One of the things I want to do as part of the rule is to update the Revision Table with the Vault data. I want to run it as an"After Open Document" rule.  This is a minimal version of the code I would like to run. MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox....

**Code:**

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox.Show("Updated Revision Table")
```

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox.Show("Updated Revision Table")
```

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)
```

```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)
```

---

## Ilogic split a string

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-split-a-string/td-p/8777436)

**Author:** Anonymous

**Date:** ‎05-08-2019
	
		
		12:20 AM

**Description:** Goodday,
I am searching for a way to split strings. Where:
Dim Separators() As Char = {";"} 
Doesnt seem to work at my code.
 
As you see in the code the separator is-> ;
for example, if the string is 1;2, it must be separated into 1 and 2 given to 2 parameters,
so spot1 -> "1" and spot2 -> "2"
This also applies to, for example, 4;6;13 in 4, 6 and 13 given to 3 parameters, etc.
 
I hope i can find the solution im looking for over here!
Kind Regards,
Thomas
					
				
			
			
				
			
			
				
	
...

**Code:**

```vb
Dim Separators() As Char = {";"}
```

```vb
Dim Separators() As Char = {";"} 
Sentence = "1;2;3;4;5"
Words = Sentence.Split(Separators)
i = 0
For Each wrd In Words
MessageBox.Show("Word Index #" & i & " = " & Words(i))
i=i+1
Next
```

```vb
Dim Separators() As Char = {";"} 
Sentence = "10;20;3;4;5"
Words = Sentence.Split(Separators)

spot1 = Words(0)
spot2 = Words(1)

Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
```

```vb
Dim Separators() As Char = {";"} 
Sentence = StringParameter
If Sentence = "1;1" Then
Words = Sentence.Split(Separators)
spot1 = Words(0)
spot2 = Words(1)
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
Else If Sentence = "1;5;7" Then
Words = Sentence.Split(Separators)
spot1 = Words(0)
spot2 = Words(1)
spot3 = Words(2)
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
End If 
InventorVb.DocumentUpdate()
```

```vb
Parameter.UpdateAfterChange = True
iLogicVb.UpdateWhenDone = True
```

```vb
Dim Separators() As Char = {";",vbCrLf} Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(Separators)
i = 0
For Each wrd In Words
MessageBox.Show("Word Index #" & i & " = " & Words(i))
i=i+1
Next
```

```vb
Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(vbCrLf)

For Each wrd In Words
	'strip out carraige returns & linefeeds
	wrd = wrd.Replace(vbCr, "").Replace(vbLf, "")
	MessageBox.Show("Word Index #" & i & " = " & wrd)
Next
```

```vb
Sentence = "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & "5"
Words = Sentence.Split(vbCrLf)

For Each wrd In Words
	'strip out carraige returns & linefeeds
	wrd = wrd.Replace(vbCr, "").Replace(vbLf, "")
	MessageBox.Show("Word Index #" & i & " = " & wrd)
Next
```

```vb
Sentence = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
```

```vb
Sentence = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
```

```vb
Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Separators() As Char = {".", "-", " " }
Dim Words() As String = Sentence.Split(Separators)
Dim Word1 As String = Words(0)
Dim Word2 As String = Words(1)
Dim Word3 As String = Words(2)
Dim Word4 As String = Words(3)
Dim Word5 As String = Words(4)
Dim Word6 As String = Words(5)
MsgBox("Word1 = " & Word1 & _
vbCrLf & "Word2 = " & Word2 & _
vbCrLf & "Word3 = " & Word3 & _
vbCrLf & "Word4 = " & Word4 & _
vbCrLf & "Word5 = " & Word5 & _
vbCrLf & "Word6 = " & Word6, vbInformation, "Sentence Split Into Words")
```

```vb
Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Words() As String = Sentence.Split({".", "-", " " },StringSplitOptions.RemoveEmptyEntries)
Dim sResults As String
For i As Integer = 0 To Words.Length - 1
	sResults &= "Word " & i & " = " & Words(i) & vbCrLf
Next
MsgBox(sResults, vbInformation, "Sentence Split Into Words")
```

```vb
Dim Sentence As String = "XX.YY.ZZZZ-R0 PROJECT DESCRIPTION"
Dim Words() As String = Sentence.Split({".", "-", " " },StringSplitOptions.RemoveEmptyEntries)
Dim sResults As String
For i As Integer = 0 To Words.Length - 1
	sResults &= "Word " & i & " = " & Words(i) & vbCrLf
Next
MsgBox(sResults, vbInformation, "Sentence Split Into Words")
```

---

