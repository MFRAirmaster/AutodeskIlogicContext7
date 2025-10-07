' Title: How to Stop Dialog Box From Popping Up
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-stop-dialog-box-from-popping-up/td-p/13834921#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:03:34.180616

Dim drawdoc As DrawingDocument = ThisApplication.ActiveDocument
If drawdoc.ReferencedDocuments.Item(1).documenttype=DocumentTypeEnum.kAssemblyDocumentObject Then
Dim oDoc As AssemblyDocument = drawdoc.ReferencedDocuments.Item(1)
Dim oBOM As BOM = oDoc.ComponentDefinition.BOM
oBOM.StructuredViewEnabled = True
oBOM.PartsOnlyViewEnabled = True
End If