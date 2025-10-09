' Title: iLogic to extract Sheet Metal Data
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-extract-sheet-metal-data/td-p/13789124
' Category: advanced
' Scraped: 2025-10-09T09:05:50.868524

Dim doc As PartDocument = ThisDoc.Document
Dim def As SheetMetalComponentDefinition = doc.ComponentDefinition
MsgBox(def.FlatPattern.TopFace.Evaluator.Area & "cm^2")