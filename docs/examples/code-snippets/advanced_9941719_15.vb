' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719
' Category: advanced
' Scraped: 2025-10-07T13:17:40.528259

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MsgBox("This rule only works on Assembly files.", vbCritical, "Wrong Document Type")
	Return
End If
Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument