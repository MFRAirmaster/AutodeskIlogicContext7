' Title: Stop an iLogic rule while it's running
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/stop-an-ilogic-rule-while-it-s-running/td-p/12304171#messageview_0
' Category: api
' Scraped: 2025-10-09T09:02:34.903921

Public Sub Main
    progBar = New TestProgressBar(ThisApplication)
    progBar.Start()
End Sub

Public Class TestProgressBar
    Private WithEvents progBar As Inventor.ProgressBar
    Private invApp As Inventor.Application

    Public Sub New(InventorApp As Inventor.Application)
        invApp = InventorApp
    End Sub

    Public Sub Start()
        progBar = invApp.CreateProgressBar(False, 10, "Test of Progress Bar", True)

        Dim j As Integer
        For j = 0 To 10
            progBar.Message = ("Current Index: " & j & "/" & 10)
            Threading.Thread.Sleep(1000)
            progBar.UpdateProgress()
        Next

        progBar.Close()
    End Sub

    Private Sub progBar_OnCancel() Handles progBar.OnCancel
        progBar.Close()
        MsgBox("Cancelled")
    End Sub
End Class