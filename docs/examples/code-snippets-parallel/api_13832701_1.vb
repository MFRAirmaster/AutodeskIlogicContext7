' Title: How to &quot;Show All&quot; in a part document with iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-show-all-quot-in-a-part-document-with-ilogic/td-p/13832701
' Category: api
' Scraped: 2025-10-07T14:02:09.383898

Dim partDoc As PartDocument
partDoc = ThisApplication.ActiveDocument

Dim compDef As PartComponentDefinition
compDef = partDoc.ComponentDefinition

Dim activeViewRep As DesignViewRepresentation = compDef.RepresentationsManager.ActiveDesignViewRepresentation
activeViewRep.ShowAll