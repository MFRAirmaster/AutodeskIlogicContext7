' Title: How to create a new assembly from ZeroDoc environment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-new-assembly-from-zerodoc-environment/td-p/13746786
' Category: advanced
' Scraped: 2025-10-09T09:00:28.105689

If Globals.g_inventorApplication Is Nothing Then
            System.Windows.Forms.MessageBox.Show("Inventor is not running or could not be launched.", CStr(vbCritical))
            Exit Sub
        End If

        ' Define the document type as an Assembly Document
        Dim oDocType As Inventor.DocumentTypeEnum
        oDocType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject

        ' Get the default template file for an assembly
        ' This will use the default template set in Inventor's options.
        Dim sTemplateFile As String
        sTemplateFile = Globals.g_inventorApplication.FileManager.GetTemplateFile(oDocType)

        ' Create a new assembly document
        Dim oAsmDoc As Inventor.AssemblyDocument
        oAsmDoc = Globals.g_inventorApplication.Documents.Add(oDocType, sTemplateFile, True) ' True makes it visible and active