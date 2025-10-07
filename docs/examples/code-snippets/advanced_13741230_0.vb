' Title: Update Inventor internal naming of components
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-inventor-internal-naming-of-components/td-p/13741230
' Category: advanced
' Scraped: 2025-10-07T13:29:20.653555

For Each oRefDoc As Document In ThisDoc.Document.AllReferencedDocuments
	oRefDoc.Dirty = True
	oRefDoc.Save
Next