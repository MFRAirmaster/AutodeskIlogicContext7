' Title: iProperties Update Help
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/iproperties-update-help/td-p/13808192
' Category: advanced
' Scraped: 2025-10-07T14:10:58.042315

Sub main
Dim ass As AssemblyDocument = ThisDoc.Document
For Each doc As Document In ass.AllReferencedDocuments
	iprops(doc)
Next
	iprops(ass)
End Sub
Sub iprops(doc As Document)
	' Get the current document's file path and name
Dim filePath As String
filePath = doc.FullFileName

' Extract the first 5 characters of the filename (optional)
Dim prefix As String
prefix = Mid(filePath, 1, 5)

' Ensure the filename has at least 12 characters
If Len(filePath) >= 12 Then
    ' Extract the filename without extension
    Dim fileName As String
    fileName = System.IO.Path.GetFileNameWithoutExtension(filePath)
    
    ' Extract the first 12 characters of the filename
    Dim partNumber As String
    partNumber = Mid(fileName, 1, 12)

    ' Access the iProperties of the document
    Dim oPropSet As PropertySet
    oPropSet = doc.PropertySets.Item("Design Tracking Properties")
    
    ' Check if "Part Number" property exists and update it
    If Not oPropSet.Item("Part Number") Is Nothing Then
        oPropSet.Item("Part Number").Value = partNumber
    Else
        ' If the Part Number property is not found, throw an error or handle it
        MessageBox.Show("Part Number property not found!")
    End If
End If
End Sub