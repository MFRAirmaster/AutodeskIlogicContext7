' Title: Export all Files of assembly to step files with part number as name
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-all-files-of-assembly-to-step-files-with-part-number-as/td-p/9941719
' Category: advanced
' Scraped: 2025-10-07T14:10:27.935747

'check that the active document is an assembly file
If ThisApplication.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
	MessageBox.Show("This Rule " & iLogicVb.RuleName & " only works on Assembly Files.", "WRONG DOCUMENT TYPE",MessageBoxButtons.OK,MessageBoxIcon.Error)
	Return
End If

'define the active document as an assembly file
Dim oAsmDoc As AssemblyDocument
oAsmDoc = ThisApplication.ActiveDocument
oAsmName = ThisDoc.FileName(False) 'without extension

If ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
    MessageBox.Show("Please run this rule from the assembly file.", "iLogic")
    Exit Sub
End If

'get user input
RUsure = MessageBox.Show ( _
"This will create a STEP file for all components." _
& vbLf & " " _
& vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
& vbLf & "This could take a while.", "iLogic - Batch Output STEPs ",MessageBoxButtons.YesNo)
If RUsure = vbNo Then
    Return
Else
End If

'- - - - - - - - - - - - -STEP setup - - - - - - - - - - - -
oPath = ThisDoc.Path
'get STEP target folder path
' original => oFolder = oPath & "\" & oAsmName & " STEP Files"
oFolder = oPath & "\STEP Files"
'Check for the step folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
System.IO.Directory.CreateDirectory(oFolder)
End If


'- - - - - - - - - - - - -Assembly - - - - - - - - - - - -
ThisDoc.Document.SaveAs(oFolder & "\" & oAsmName &(".stp") , True)

'- - - - - - - - - - - - -Components - - - - - - - - - - - -
'look at the files referenced by the assembly
Dim oRefDocs As DocumentsEnumerator
oRefDocs = oAsmDoc.AllReferencedDocuments
Dim oRefDoc As Document
'work the referenced models
For Each oRefDoc In oRefDocs
    Dim oCurFile As Document
    oCurFile = ThisApplication.Documents.Open(oRefDoc.FullFileName, True)
    oCurFileName = oCurFile.FullFileName
   
    'defines backslash As the subdirectory separator
    Dim strCharSep As String = System.IO.Path.DirectorySeparatorChar
   
    'find the postion of the last backslash in the path
    FNamePos = InStrRev(oCurFileName, "\", -1)  
    'get the file name with the file extension
    Name = Right(oCurFileName, Len(oCurFileName) - FNamePos)
    'get the file name (without extension)
    ShortName = Left(Name, Len(Name) - 4)

    Try
        oCurFile.SaveAs(oFolder & "\" & ShortName & (".stp") , True)
    Catch
        MessageBox.Show("Error processing " & oCurFileName, "ilogic")
    End Try
    oCurFile.Close
Next
'- - - - - - - - - - - - -
MessageBox.Show("New Files Created in: " & vbLf & oFolder, "iLogic")
)