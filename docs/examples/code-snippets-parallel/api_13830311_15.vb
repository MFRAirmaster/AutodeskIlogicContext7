' Title: Inventor 2026 Inventor Application, GetActiveObject is not a member of Marshal
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2026-inventor-application-getactiveobject-is-not-a/td-p/13830311#messageview_0
' Category: api
' Scraped: 2025-10-09T09:03:26.476615

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