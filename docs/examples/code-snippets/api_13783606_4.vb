' Title: Macro to delete specific sketch symbol?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606
' Category: api
' Scraped: 2025-10-07T13:12:45.818204

Public Sub DeleteSpecificSketchSymbol()
    ' Set a reference to the active drawing document.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    ' Define the name of the sketch symbol to delete.
    Const sSymbolName As String = "Rev Triangle" ' Replace with the actual symbol name
    ' Loop through each sheet in the drawing document.
    Dim oSheet As Sheet
    For Each oSheet In oDrawDoc.Sheets
        ' Activate the current sheet (optional, but good practice for visibility).
        oSheet.Activate
        ' Loop through each sketched symbol on the current sheet.
        Dim oSketchedSymbol As SketchedSymbol
        Dim i As Long
        For i = oSheet.SketchedSymbols.Count To 1 Step -1 ' Loop backwards to avoid issues when deleting
            Set oSketchedSymbol = oSheet.SketchedSymbols.Item(i)
            ' Check if the symbol's name matches the target name.
            If oSketchedSymbol.Name = sSymbolName Then
                ' Delete the instance of the sketch symbol.
                oSketchedSymbol.Delete
            End If
        Next i
    Next oSheet
End Sub