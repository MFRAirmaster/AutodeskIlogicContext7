' Title: iLogic for hide/show every possible object
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-for-hide-show-every-possible-object/td-p/13761985#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:05:30.688902

Sub Main()
' iLogic: Sichtbarkeit umschalten je nach Dokumenttyp und Auswahl

Dim oDoc As Document = ThisApplication.ActiveDocument
Dim oSelSet As SelectSet = oDoc.SelectSet

If oSelSet.Count = 0 Then
    MsgBox("Bitte mindestens ein Objekt auswählen.", vbExclamation)
    Return
End If

Select Case oDoc.DocumentType

    '------------------------------
    ' PART-Dokument (.ipt)
    '------------------------------
    Case DocumentTypeEnum.kPartDocumentObject
        For Each oObj In oSelSet
            Call ToggleVisibility(oObj)
        Next

    '------------------------------
    ' ASSEMBLY-Dokument (.iam)
    '------------------------------
    Case DocumentTypeEnum.kAssemblyDocumentObject
        For Each oObj In oSelSet
            Call ToggleVisibility(oObj)
        Next

    '------------------------------
    ' DRAWING-Dokument (.idw)
    '------------------------------
    Case DocumentTypeEnum.kDrawingDocumentObject
        For Each oObj In oSelSet
            Call ToggleVisibility(oObj)
        Next

    Case Else
        MsgBox("Dokumenttyp wird nicht unterstützt.", vbExclamation)
End Select

End Sub

'--------------------------------------
' Hilfsfunktion: Sichtbarkeit umschalten
'--------------------------------------
Sub ToggleVisibility(oObj As Object)
    Try
        ' Falls das Objekt die Property "Visible" hat
        Dim t As Type = oObj.GetType()
        Dim p = t.GetProperty("Visible")
        If Not p Is Nothing Then
            Dim current As Boolean = CBool(p.GetValue(oObj, Nothing))
            p.SetValue(oObj, Not current, Nothing)
            Exit Sub
        End If

        ' Spezielle Fälle behandeln:
        If TypeOf oObj Is ComponentOccurrence Then
            oObj.Visible = Not oObj.Visible
        ElseIf TypeOf oObj Is Sketch Then
            oObj.Visible = Not oObj.Visible
        ElseIf TypeOf oObj Is WorkAxis Then
            oObj.Visible = Not oObj.Visible
        ElseIf TypeOf oObj Is WorkPlane Then
            oObj.Visible = Not oObj.Visible
        ElseIf TypeOf oObj Is WorkPoint Then
            oObj.Visible = Not oObj.Visible
        ElseIf TypeOf oObj Is DrawingCurveSegment Then
            oObj.Visible = Not oObj.Visible
		ElseIf TypeOf oObj Is DrawingCurve Then
            Dim oCurve As DrawingCurve = oObj
            Dim oOcc As ComponentOccurrence = oCurve.ModelGeometry.ContainingOccurrence
            If Not oOcc Is Nothing Then
                Dim oView As DrawingView = oCurve.Parent
                oView.SetVisibility(oOcc, Not oView.GetVisibility(oOcc))
            End If
        Else
            ' Optional: Debug-Ausgabe für nicht unterstützte Typen
            ' MsgBox("Kein Sichtbarkeits-Flag für: " & oObj.ToString)
        End If

    Catch ex As Exception
        MsgBox("Fehler bei: " & oObj.ToString & vbCrLf & ex.Message)
    End Try
End Sub