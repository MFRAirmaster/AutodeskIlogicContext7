' Title: Browser/Tree doesn't update to reflect visibility tag when changed by rule
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/browser-tree-doesn-t-update-to-reflect-visibility-tag-when/td-p/13744987#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:51:44.868974

Dim oDoc As Document
Dim Occ As ComponentOccurrence
Dim comp As ComponentOccurrencesEnumerator
Dim invApp As Inventor.Application
invApp = ThisApplication
oDoc = ThisDoc.Document 'local variable to shorthand to current document
comp =oDoc.ComponentDefinition.Occurrences.AllleafOccurrences '"AllLeafOccurrences" make rule read the lowest level components making hidden parts within subassemblies visible (unclear of top-level equivalent)
For Each Occ In comp 'Applies following to all components encompased
   Occ.Visible = True 'Sets visible boolean to True
Next