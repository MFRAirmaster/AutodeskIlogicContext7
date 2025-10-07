' Title: Change Sketch Symbol by Ilogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/change-sketch-symbol-by-ilogic/td-p/13742729
' Category: api
' Scraped: 2025-10-07T13:10:48.793703

' Obter o documento de desenho atual
Dim oDrawDoc As DrawingDocument = ThisApplication.ActiveDocument
Dim oSheet As Sheet = oDrawDoc.Sheets(1)

' Percorrer todos os símbolos da folha
For Each oSymbol As SketchedSymbol In oSheet.SketchedSymbols
    If oSymbol.Definition.Name = "PA" Then
        ' Substitui o valor do prompt chamado "OBSERVAÇÕES"	
		Dim oTextBoxes As TextBoxes = oSymbol.Definition.Sketch.TextBoxes
		For Each oTBox As Inventor.TextBox In oTextBoxes
			If Not oTBox.Text = "OBSERVAÇÕES" Then Continue For
			If oSymbol.GetResultText(oTBox) = "PERFIL ALTO" Then
				oSymbol.SetPromptResultText(oTBox, "HIGH PROFILE")
				Exit For
			End If
		Next
    End If
Next