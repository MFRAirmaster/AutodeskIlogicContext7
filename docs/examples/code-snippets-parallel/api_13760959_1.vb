' Title: Show Thumbnail in the Windows.Forms.PictureBox iLogic Inventor 2026
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/show-thumbnail-in-the-windows-forms-picturebox-ilogic-inventor/td-p/13760959
' Category: api
' Scraped: 2025-10-07T14:08:22.314915

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