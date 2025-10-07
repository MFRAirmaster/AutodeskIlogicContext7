' Title: Run external rule without an open document?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/run-external-rule-without-an-open-document/td-p/10149441#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:10:30.958138

If ThisApplication.ActiveDocument Is Nothing Then
ThisApplication.Documents.Add kPartDocumentObject
iLogicAuto.RunExternalRule ThisApplication.ActiveDocument, oPath
ThisApplication.Documents.VisibleDocuments.Item(1).Close (True)
Else
iLogicAuto.RunExternalRule ThisApplication.ActiveDocument, oPath
End If