' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:12.526503

'check that the active document is an assembly file
If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
    MessageBox.Show("This Rule " & iLogicVb.RuleName & " only works on Assembly Files.", "WRONG DOCUMENT TYPE", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Return
End If

'define the active document as an assembly file
Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim oAsmName As String = ThisDoc.FileName(False) 'without extension


'get user input
Dim RUsure = MessageBox.Show(
"This will create a STEP file for all components." _
& vbLf & " " _
& vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
& vbLf & "This could take a while.", "iLogic - Batch Output STEPs ", MessageBoxButtons.YesNo)
If RUsure = vbNo Then
    Return
Else
End If

'- - - - - - - - - - - - -STEP setup - - - - - - - - - - - -
Dim oPath = ThisDoc.Path
'get STEP target folder path
' original => oFolder = oPath & "\" & oAsmName & " STEP Files"
Dim oFolder = oPath & "\STEP Files"
'Check for the step folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
    System.IO.Directory.CreateDirectory(oFolder)
End If


'- - - - - - - - - - - - -Assembly - - - - - - - - - - - -
ThisDoc.Document.SaveAs(oFolder & "\" & oAsmName & (".stp"), True)

'- - - - - - - - - - - - -Components - - - - - - - - - - - -
'look at the files referenced by the assembly
Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments
'work the referenced models
For Each oRefDoc As Document In oRefDocs
    Dim oCurFile As Document = ThisApplication.Documents.Open(oRefDoc.FullFileName, True)
    Dim oCurFileName = oCurFile.FullFileName
    Dim ShortName = IO.Path.GetFileNameWithoutExtension(oCurFileName)

    Dim oPropSets As PropertySets = oCurFile.PropertySets
    Dim oPropSet As PropertySet = oPropSets.Item("Design Tracking Properties")
    Dim oPartNumiProp As [Property] = oPropSet.Item("Part Number")


    Try
        oCurFile.SaveAs(oFolder & "\" & oPartNumiProp.Value & (".stp"), True)
    Catch
        MessageBox.Show("Error processing " & oCurFileName, "ilogic")
    End Try
    oCurFile.Close()
Next
'- - - - - - - - - - - - -
MessageBox.Show("New Files Created in: " & vbLf & oFolder, "iLogic")