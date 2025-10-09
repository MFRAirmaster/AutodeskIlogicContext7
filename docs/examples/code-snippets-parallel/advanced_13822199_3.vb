' Title: Help on API AddByiMateAndEntity
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/help-on-api-addbyimateandentity/td-p/13822199
' Category: advanced
' Scraped: 2025-10-09T08:50:31.486024

Dim imdp as object
occ.CreateGeometryProxy(oImate , imdp) ' proxy form occurrence containing imatedefinition
Dim oiMateResult As iMateResult = ass.ComponentDefinition.iMateResults.AddByiMateAndEntity(imdp , oface)