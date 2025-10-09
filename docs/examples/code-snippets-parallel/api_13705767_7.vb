' Title: iLogic Code to export Sheets to IDWS
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-export-sheets-to-idws/td-p/13705767#messageview_0
' Category: api
' Scraped: 2025-10-09T09:05:14.098880

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