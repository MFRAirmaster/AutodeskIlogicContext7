' Title: Problem to read open documents
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/problem-to-read-open-documents/td-p/13825440#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:56:36.661457

Private Function GetDocumentList(ByRef docs As Object) As Variant

    Dim assiemiCount As Integer
    Dim i As Integer
    Dim doc As Object
    Dim docNames, selectedDocName As String
    Dim isAssemblyDocument, selezioneValida, gestioneCiclo As Boolean
    
    For i = 1 To docs.VisibleDocuments.Count
        Set doc = docs.Item(i)
        isAssemblyDocument = TypeName(doc) = "AssemblyDocument"
           
        If InStr(doc.FullFileName, "iam") And isAssemblyDocument Then
            docNames = docNames & i & ": " & doc.DisplayName & vbCrLf
            assiemiCount = assiemiCount + 1
        End If

    Next i
    
    If assiemiCount ....