' Title: USE ILOGIC TO EXPORT FLAT PATTERN USING INI TEMPLATE FILE
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/use-ilogic-to-export-flat-pattern-using-ini-template-file/td-p/13765963#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:01:22.241263

'Set a reference to the active document (the document to be published).
    Dim oDoc As PartDocument
    oDoc = ThisApplication.ActiveDocument
    
    Dim oCompDef As SheetMetalComponentDefinition
   oCompDef = oDoc.ComponentDefinition
        
    If oCompDef.HasFlatPattern = False Then
        oCompDef.Unfold
    End If
    
    Dim oFlatPattern As FlatPattern
    oFlatPattern = oCompDef.FlatPattern
    
	' Get the DWG translator Add-In.
    Dim DWGAddIn As TranslatorAddIn
   DWGAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
    
    Dim oContext As TranslationContext
    oContext = ThisApplication.TransientObjects.CreateTranslationContext
	oContext.Type = kFileBrowseIOMechanism
  
  ' Create a NameValueMap object
    Dim oOptions As NameValueMap
    oOptions = ThisApplication.TransientObjects.CreateNameValueMap
           
    ' Create a DataMedium object
        Dim oData As DataMedium
        oData = ThisApplication.TransientObjects.CreateDataMedium
		
	'Set the destination file name
        oData.FileName = oDoc.FullFileName & ".dwg"
		
	' Check whether the translator has 'SaveCopyAs' options
    If DWGAddIn.HasSaveCopyAsOptions(oFlatPattern, oContext, oOptions) Then
	'Set ini file for export options
        Dim strIniFile As String
        strIniFile = "C:\tempDWGOut.ini"
        ' Create the name-value that specifies the ini file to use.
        oOptions.Value("Export_Acad_IniFile") = strIniFile
    End If
	
	
    'Publish document.
	Call DWGAddIn.SaveCopyAs(oFlatPattern, oContext, oOptions, oData)