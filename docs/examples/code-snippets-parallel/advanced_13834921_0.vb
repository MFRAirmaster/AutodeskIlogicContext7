' Title: How to Stop Dialog Box From Popping Up
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-stop-dialog-box-from-popping-up/td-p/13834921#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:03:34.180616

Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim oBOM As BOM = oDoc.ComponentDefinition.BOM
oBOM.StructuredViewEnabled = True
oBOM.PartsOnlyViewEnabled = True