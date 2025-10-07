' Title: LEADER NOTE FILTER SETTINGS
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/leader-note-filter-settings/td-p/13786343
' Category: advanced
' Scraped: 2025-10-07T13:32:24.940625

Dim oDNote As Inventor.DrawingNote = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingNoteFilter, "Select a drawing note to inspect.")
If (oDNote Is Nothing) Then Return
Logger.Info(vbCrLf & oDNote.FormattedText)
Dim sFText As String = InputBox("Here is the LeaderNote.FormattedText:", "LeaderNote.FormattedText", oDNote.FormattedText)
oDNote.FormattedText = sFText