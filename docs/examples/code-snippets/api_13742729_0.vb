' Title: Change Sketch Symbol by Ilogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/change-sketch-symbol-by-ilogic/td-p/13742729
' Category: api
' Scraped: 2025-10-07T13:10:48.793703

' Obter o documento de desenho atual
Dim oDrawDoc As DrawingDocument = ThisApplication.ActiveDocument
Dim oSheet As Sheet = oDrawDoc.Sheets(1)

' Percorrer todos os símbolos da folha
Dim oSymbol
For Each oSymbol In oSheet.SketchedSymbols
    If oSymbol.Definition.Name = "PA" Then
        ' Substitui o valor do prompt chamado "OBSERVAÇÕES"
        oSymbol.SetPromptResultText("OBSERVAÇÕES", "HIGH PROFILE")
    End If
Next