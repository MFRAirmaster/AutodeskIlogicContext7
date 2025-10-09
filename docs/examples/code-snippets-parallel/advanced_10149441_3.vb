' Title: Run external rule without an open document?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/run-external-rule-without-an-open-document/td-p/10149441
' Category: advanced
' Scraped: 2025-10-09T09:02:52.221770

If ThisApplication.ActiveDocument Is Nothing Then
ThisApplication.Documents.Add kPartDocumentObject
iLogicAuto.RunExternalRule ThisApplication.ActiveDocument, oPath
ThisApplication.Documents.VisibleDocuments.Item(1).Close (True)
Else
iLogicAuto.RunExternalRule ThisApplication.ActiveDocument, oPath
End If