' Title: Unreserve legacy check out file with code?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/unreserve-legacy-check-out-file-with-code/td-p/12232593
' Category: advanced
' Scraped: 2025-10-07T14:19:21.609875

Sub Main removeCheckOutStatusSample()
    ' to remove the Check out status you need to activate a Shared Project
    ' then open documents and call below function, and then save the documents.
    Dim oDoc As Document
    For Each oDoc In ThisApplication.Documents
        removeCheckOutStatus (oDoc)
    Next
    ThisApplication.Documents.CloseAll (False)
End Sub
Sub removeCheckOutStatus(oDoc As Document)
    If oDoc.ReservedForWriteName <> "" Then
        oDoc.ReservedForWriteByMe = True
        oDoc.RevertReservedForWriteByMe
        Dim oSaveCmd As ControlDefinition
        oSaveCmd = ThisApplication.CommandManager.ControlDefinitions("AppFileSaveCmd")
        oSaveCmd.Execute
    End If
End Su