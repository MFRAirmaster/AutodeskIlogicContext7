' Title: Defining A-Side for sheet metal part using Entitys
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/defining-a-side-for-sheet-metal-part-using-entitys/td-p/13759776#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:10:42.569110

Dim oNEs As NamedEntities = iLogicVb.Automation.GetNamedEntities(oDoc)
Dim oMyNamedEntity = oNEs.TryGetEntity("entity name")