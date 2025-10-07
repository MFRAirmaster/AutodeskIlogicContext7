' Title: problem with rules on remote pc
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/problem-with-rules-on-remote-pc/td-p/13810727#messageview_0
' Category: api
' Scraped: 2025-10-07T14:16:02.996377

Imports System.IO

Sub Main()
' Define source folder
Dim sourceFolder As String = "C:\Users\Frank\Documents\Inventor\ilogic Afsluiter"

' Ask user to choose destination folder
Dim fbd As New FolderBrowserDialog()
fbd.Description = "Select the destination folder"
fbd.ShowNewFolderButton = True

Dim result As DialogResult = fbd.ShowDialog()
If result <> DialogResult.OK Then
MessageBox.Show("Operation cancelled.", "iLogic")
Exit Sub
End If

Dim destFolder As String = fbd.SelectedPath

' Ask user for prefix (e.g. M593)
Dim prefix As String = InputBox("Enter the prefix (example: M593)", "File Prefix", "M593")
If prefix = "" Then
MessageBox.Show("No prefix entered. Operation cancelled.", "iLogic")
Exit Sub
End If

' Create destination folder if it doesn’t exist
If Not Directory.Exists(destFolder) Then
Directory.CreateDirectory(destFolder)
End If

' Get all files in source folder
Dim files() As String = Directory.GetFiles(sourceFolder)

' Counter for renaming
Dim counter As Integer = 0

For Each filePath As String In files
Dim ext As String = Path.GetExtension(filePath)
Dim newName As String = prefix & "-" & counter.ToString("000") & "-00" & ext
Dim destPath As String = Path.Combine(destFolder, newName)

' Copy file (overwrite if already exists)
File.Copy(filePath, destPath, True)

counter += 1
Next

MessageBox.Show("Files copied and renamed successfully to:" & vbCrLf & destFolder, "iLogic")
End Sub