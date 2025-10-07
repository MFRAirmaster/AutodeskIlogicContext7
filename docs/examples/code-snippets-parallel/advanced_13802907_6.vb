' Title: Save and Replace Rule
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/save-and-replace-rule/td-p/13802907#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:05:21.712644

Option Explicit On
Sub Main()
	
	Dim prefix As String = "25100"
	Dim fileName As String = System.IO.Path.GetFileName(ThisDoc.Document.FullFileName)
	Dim newFileName As String = fileName.Replace("PROJECT", prefix)
	If newFileName <> fileName Then ThisDoc.Document.SaveAs(ThisDoc.Path & "\" & newFileName, False)
	
	CopyAndReplace(ThisDoc.Document, prefix)
	
End Sub 'Main

Sub CopyAndReplace(doc As AssemblyDocument, prefix As String)
	
	Dim fileName As String
	Dim newFileName As String
	Dim suppressed As Boolean
	For Each comp As ComponentOccurrence In doc.ComponentDefinition.Occurrences
		fileName = GetFileName(comp)
		If fileName Is Nothing Then Continue For
		If Not System.IO.Path.GetFileName(fileName).StartsWith("PROJECT") Then
			Logger.Info("Skipping: " & comp.Name)
			Continue For
		Else
			Logger.Info("Copying: " & comp.Name)
		End If
		newFileName = fileName.Replace("PROJECT", prefix)
		
		suppressed = comp.Suppressed
		If suppressed Then comp.Unsuppress
			
		CopyFile(fileName, newFileName)
		comp.Replace(newFileName, True)
		
		If TypeOf comp.Definition Is AssemblyComponentDefinition Then CopyAndReplace(comp.Definition.Document, prefix)
		
		If suppressed Then comp.Suppress
	Next comp

End Sub 'CopyAndReplace

Sub CopyFile(sourceFile As String, targetFile As String)

    If Not System.IO.File.Exists(targetFile) Then System.IO.File.Copy(sourceFile, targetFile, True)
    My.Computer.FileSystem.GetFileInfo(targetFile).IsReadOnly = False
    
End Sub 'CopyFile

Function GetFileName(comp As ComponentOccurrence) As String

    Dim fileName As String = Nothing

    If comp IsNot Nothing AndAlso comp.ReferencedFileDescriptor IsNot Nothing Then
        fileName = comp.ReferencedFileDescriptor.FullFileName
    End If

    Return fileName

End Function 'GetFilename