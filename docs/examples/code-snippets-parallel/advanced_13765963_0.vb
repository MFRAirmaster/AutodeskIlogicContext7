' Title: USE ILOGIC TO EXPORT FLAT PATTERN USING INI TEMPLATE FILE
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/use-ilogic-to-export-flat-pattern-using-ini-template-file/td-p/13765963#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:01:22.241263

' Get the DWG translator Add-In.
    Dim DWGAddIn As TranslatorAddIn
   DWGAddIn = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")

    'Set a reference to the active document (the document to be published).
    Dim oDocument As Document
    oDocument = ThisApplication.ActiveDocument

    Dim oContext As TranslationContext
    oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = kFileBrowseIOMechanism

    ' Create a NameValueMap object
    Dim oOptions As NameValueMap
    oOptions = ThisApplication.TransientObjects.CreateNameValueMap

    ' Create a DataMedium object
    Dim oDataMedium As DataMedium
    oDataMedium = ThisApplication.TransientObjects.CreateDataMedium

    ' Check whether the translator has 'SaveCopyAs' options
    If DWGAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

        Dim strIniFile As String
        strIniFile = "C:\tempDWGOut.ini"
        ' Create the name-value that specifies the ini file to use.
        oOptions.Value("Export_Acad_IniFile") = strIniFile
    End If

    'Set the destination file name
    oDataMedium.FileName = "c:\tempdwgout.dwg"

    'Publish document.
    Call DWGAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)