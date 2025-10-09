' Title: DXF method with INI not working, sOut version works
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/dxf-method-with-ini-not-working-sout-version-works/td-p/13844011
' Category: advanced
' Scraped: 2025-10-09T09:05:10.645616

Sub Main

	DXFFolder = "C:\Temp\DXF\"
	oModelDocName = ThisDoc.Document.FullFileName
	DXFName = IO.Path.GetFileNameWithoutExtension(oModelDocName) & "-INI"

	Dim oInv As Application = ThisApplication
	Dim oDoc As Document = oInv.Documents.Open(oModelDocName, True)

	Dim oSheetDef As SheetMetalComponentDefinition = oDoc.ComponentDefinition

	If oSheetDef.HasFlatPattern = False Then
		oSheetDef.Unfold
	Else
		oSheetDef.FlatPattern.Edit
	End If

	'Set the DXF target file name
	If Not IO.Directory.Exists(DXFFolder) Then IO.Directory.CreateDirectory(DXFFolder)

	'Get the DXF translator Add-In.
	Dim DXFAddIn As TranslatorAddIn
	DXFAddIn = oInv.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")

	'Create a DataMedium object
	Dim oDataMedium As DataMedium
	oDataMedium = oInv.TransientObjects.CreateDataMedium

	'Create a TranslationContext object
	Dim oContext As TranslationContext
	oContext = ThisApplication.TransientObjects.CreateTranslationContext
	oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

	'Get the Design Data Path
	Dim INI_Folder As String
	INI_Folder = "C:\Temp\DXF\DXF configs\"

	'Create a NameValueMap object
	Dim oOptions As NameValueMap
	oOptions = oInv.TransientObjects.CreateNameValueMap

	'Specify the INI file path for exporting
	Dim iniFilePath As String = INI_Folder & "PC_flatpatterns.ini"
	oOptions.Value("Export_Acad_IniFile") = iniFilePath

	'MsgBox(INI_Folder & "PC_flatpatterns.ini")

	'Define the file name
	oDataMedium.FileName = DXFFolder & DXFName & ".dxf"
	
		MsgBox(	oDataMedium.FileName )


	'	'Specify the AutoCAD version to export to (AutoCAD 2004)
	'	oOptions.Value("Export_AcadVersion") = 16

	'Publish document.
	DXFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)



End Sub