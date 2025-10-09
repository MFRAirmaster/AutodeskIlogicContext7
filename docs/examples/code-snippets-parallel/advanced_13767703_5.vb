' Title: Issue with Vault revision table
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/issue-with-vault-revision-table/td-p/13767703#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:38.123295

Dim oDoc As DrawingDocument = app.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    If oDoc.DocumentType <> kDrawingDocumentObject Then Exit Sub