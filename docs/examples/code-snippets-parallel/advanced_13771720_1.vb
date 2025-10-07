' Title: iAssembly iLogic to overwrite weight
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/iassembly-ilogic-to-overwrite-weight/td-p/13771720
' Category: advanced
' Scraped: 2025-10-07T14:10:57.574269

Dim oAsmDoc As AssemblyDocument
oAsmDoc = ThisApplication.ActiveDocument

Dim totalWeight As Double = 0

For Each oOcc As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
    Try
        ' Prüfen, ob Occurrence ein iAssembly-Mitglied ist
        If oOcc.IsiAssemblyMember Then
            ' iAssemblyMember holen
            Dim member As iAssemblyMember = oOcc.Definition.iAssemblyMember

            ' Aktuelle Zeile
            Dim row As iAssemblyTableRow = member.Row

            If Not row Is Nothing Then
                ' Zelle "Masse" auslesen
                Dim cell As iAssemblyTableCell = row.Item("Masse")
                Dim variantWeight As Double = CDbl(cell.Value)

                totalWeight += variantWeight
            End If
			 Else
            ' Normales Teil / Baugruppe → Inventor-Masse
            Dim occMass As Double = oOcc.MassProperties.Mass
            totalWeight += occMass
        End If
    Catch ex As Exception
        ' suppressed Occurrences oder fehlende Spalten überspringen
    End Try
Next

' Physikalische Masse überschreiben
Dim massProps As MassProperties
massProps = oAsmDoc.ComponentDefinition.MassProperties

oAsmDoc.ComponentDefinition.MassProperties.Mass = totalWeight

' In iProperties schreiben
iProperties.Value("Custom", "Masse") = totalWeight

MessageBox.Show("Masse = " & totalWeight & " kg", "iLogic Info")