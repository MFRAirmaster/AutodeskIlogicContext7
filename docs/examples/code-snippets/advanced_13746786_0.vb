' Title: How to create a new assembly from ZeroDoc environment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-new-assembly-from-zerodoc-environment/td-p/13746786
' Category: advanced
' Scraped: 2025-10-07T13:14:43.568208

'This is the code that does the real work when your command is executed.
    Sub Run_CommandCode()

        Debugger.Break()
        MsgBox("This command creates a new assembly from the ZeroDoc environment.")

        ' Get the Inventor Application object
        Dim oApp As Inventor.Application
        oApp = ThisApplication ' If running within Inventor (VBA macro)

        Dim invApp As Inventor.Application
        Try
           invApp = GetObject(, "Inventor.Application")
           '            invApp = ThisApplication ' If running within Inventor (VBA macro)
        Catch ex As Exception
           invApp = CreateObject("Inventor.Application")
           invApp.Visible = True
        End Try

        If invApp Is Nothing Then
           System.Windows.Forms.MessageBox.Show("Inventor is not running or could not be launched.", CStr(vbCritical))
           Exit Sub
        End If

        ' Define the document type as an Assembly Document
        Dim oDocType As Inventor.DocumentTypeEnum
        oDocType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject

        ' Get the default template file for an assembly
        ' This will use the default template set in Inventor's options.
        Dim sTemplateFile As String
        sTemplateFile = invApp.FileManager.GetTemplateFile(oDocType)

        ' Create a new assembly document
        Dim oAsmDoc As Inventor.AssemblyDocument
        oAsmDoc = invApp.Documents.Add(oDocType, sTemplateFile, True) ' True makes it visible and active

    End Sub