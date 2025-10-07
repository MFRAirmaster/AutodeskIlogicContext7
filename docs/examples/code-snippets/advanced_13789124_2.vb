' Title: iLogic to extract Sheet Metal Data
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-extract-sheet-metal-data/td-p/13789124#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:37:34.651599

Dim doc As PartDocument = ThisDoc.Document
Dim FlatArea As SheetMetalComponentDefinition = doc.ComponentDefinition

Parameter("FlatArea") = FlatArea.FlatPattern.TopFace.Evaluator.Area / 6.4516