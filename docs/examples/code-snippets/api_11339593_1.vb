' Title: Inventor DWG to AutoCAD DWG Conversion
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inventor-dwg-to-autocad-dwg-conversion/td-p/11339593
' Category: api
' Scraped: 2025-10-07T12:30:20.812121

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