' Title: How to check if a file has been saved in Inventor 2022?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-check-if-a-file-has-been-saved-in-inventor-2022/td-p/10738478
' Category: advanced
' Scraped: 2025-10-09T08:50:29.041962

Sub Main()
	
	Debugger.Break
	
    ' Ensure we're in an assembly document
    If ThisDoc.Document.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("Please run this rule from an assembly document.", "iLogic")
        Return
    End If

    ' Get the active assembly document
    Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument

	' Report counter and dirty status
	MessageBox.Show( _
		"FileSaveCounter = " & oAsmDoc.FileSaveCounter & vbCr & _
		"Dirty: " & oAsmDoc.Dirty & vbCr & _
		" ",
		"REPORT",
		MessageBoxButtons.OK,
		MessageBoxIcon.Information
		)
	
End Sub