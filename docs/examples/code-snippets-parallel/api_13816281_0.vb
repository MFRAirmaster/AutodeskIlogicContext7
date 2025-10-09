' Title: Copy Design Issue – Suppressed Parts Not Rebuilding or Updating (Skeleton Reference Problem)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/copy-design-issue-suppressed-parts-not-rebuilding-or-updating/td-p/13816281
' Category: api
' Scraped: 2025-10-09T09:09:31.660015

'This Rule Helps to Update all parts and assemblies (Even Suppressed) and Save.

For Each oDoc As Document In ThisApplication.Documents
Try
' Force update
oDoc.Update()

' Save after update
If oDoc.FullFileName <> "" AndAlso oDoc.IsModifiable Then
oDoc.Save2(True) ' True = Save silently without dialogs
End If
Catch
End Try
Next
iLogicVb.UpdateWhenDone = True