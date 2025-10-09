' Title: iAssembly iLogic to overwrite weight
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/iassembly-ilogic-to-overwrite-weight/td-p/13771720#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:49:11.552630

Dim oAsmDoc As AssemblyDocument
oAsmDoc = ThisApplication.ActiveDocument

Dim totalWeight As Double = 0

For Each oOcc As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
    Try
        Dim occDoc As Document = oOcc.Definition.Document

        If occDoc.DocumentType = kPartDocumentObject Or occDoc.DocumentType = kAssemblyDocumentObject Then
            If occDoc.ComponentDefinition.IsiAssemblyMember Then
                Dim member As iAssemblyMember = occDoc.ComponentDefinition.iAssemblyMember

                ' Die aktuell in der Baugruppe verwendete Variante holen
                Dim row As iAssemblyTableRow = member.ActiveTableRow

                If Not row Is Nothing Then
                  
                    Dim variantWeight As Double = CDbl(row("Masse").Value)
                    totalWeight += variantWeight
                End If
            End If
        End If
    Catch ex As Exception
    End Try
Next


' Physikalische Masse überschreiben
oAsmDoc.ComponentDefinition.MassProperties.Mass = totalWeight