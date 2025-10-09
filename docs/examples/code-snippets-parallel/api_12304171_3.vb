' Title: Stop an iLogic rule while it's running
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/stop-an-ilogic-rule-while-it-s-running/td-p/12304171#messageview_0
' Category: api
' Scraped: 2025-10-09T09:02:34.903921

' iLogic: Save all open documents, skipping any that error.
' Drops dialogs (SilentOperation) and restores settings afterward.

Dim app As Inventor.Application = ThisApplication
Dim docs As Inventor.Documents = app.Documents

Dim total As Integer = 0
Dim saved As Integer = 0
Dim failed As New System.Collections.Generic.List(Of String)

' Remember app settings
Dim originalSilent As Boolean = app.SilentOperation
Dim originalScreen As Boolean = app.ScreenUpdating

Try
    app.SilentOperation = True
    app.ScreenUpdating = False

    ' Take a snapshot of currently open docs (collection can change while saving)
    Dim openDocs As New System.Collections.Generic.List(Of Inventor.Document)
    For Each d As Inventor.Document In docs
        openDocs.Add(d)
    Next

    For Each d As Inventor.Document In openDocs
        total += 1
        Try
            d.Save()                           ' Try to save everything
            saved += 1
        Catch ex As System.Exception
            failed.Add(d.DisplayName & " — " & ex.Message)
            ' Skip and continue with the next document
        End Try
    Next

Finally
    app.SilentOperation = originalSilent
    app.ScreenUpdating = originalScreen
End Try

Dim report As String = _
    "Save All — iLogic" & vbCrLf & vbCrLf & _
    "Open docs: " & total & vbCrLf & _
    "Saved: " & saved & vbCrLf & _
    "Failed: " & failed.Count

If failed.Count > 0 Then
    report &= vbCrLf & vbCrLf & "Failures:" & vbCrLf & "— " & String.Join(vbCrLf & "— ", failed.ToArray())
End If

MsgBox(report, vbInformation, "iLogic Save All")