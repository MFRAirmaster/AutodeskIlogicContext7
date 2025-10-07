' Title: Inventor Vb.Net &quot;AddIn&quot; - Change A Drawing View
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inventor-vb-net-quot-addin-quot-change-a-drawing-view/td-p/13828621
' Category: api
' Scraped: 2025-10-07T13:55:16.646782

Sub ReplaceModelReference(drawing As DrawingDocument, newFile As String)
    'This assumes the drawing references only one model

    ' EDIT: Line replaced as suggested by @WCrihfield 
    'Dim fileDescriptor As FileDescriptor = drawing.ReferencedFileDescriptors(1).DocumentDescriptor.ReferencedFileDescriptor
    Dim fileDescriptor As FileDescriptor = drawing.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor


    fileDescriptor.ReplaceReference(newFile)
End Sub