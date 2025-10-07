' Title: acad.exe link problem!
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/acad-exe-link-problem/td-p/13822188#messageview_0
' Category: api
' Scraped: 2025-10-07T13:01:29.516705

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