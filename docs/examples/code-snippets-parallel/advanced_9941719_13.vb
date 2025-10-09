' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:12.526503

If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MessageBox.Show("This rule only works on Assembly files.", "Wrong Document Type", MessageBoxButtons.OK, MessageBoxIcon.Error)
	Return
End If
Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument