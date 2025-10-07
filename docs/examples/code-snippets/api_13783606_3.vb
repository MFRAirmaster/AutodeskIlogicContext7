' Title: Macro to delete specific sketch symbol?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606
' Category: api
' Scraped: 2025-10-07T13:12:45.818204

Sub DeleteRevTriangleSketchedSymbols()
    If Not ThisApplication.ActiveDocumentType = kDrawingDocumentObject Then Return
    Dim oDDoc As DrawingDocument
    Set oDDoc = ThisApplication.ActiveDocument
    Dim oSheet As Inventor.Sheet
    For Each oSheet In oDDoc.Sheets
        Dim oSS As SketchedSymbol
        For Each oSS In oSheet.SketchedSymbols
            If oSS.Name = "rev triangle" Then
                Call oSS.Delete
            End If
        Next 'oSS
    Next 'oSheet
End Sub