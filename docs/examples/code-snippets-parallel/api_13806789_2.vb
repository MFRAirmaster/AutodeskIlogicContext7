' Title: Decimal places in iLogic form
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/decimal-places-in-ilogic-form/td-p/13806789#messageview_0
' Category: api
' Scraped: 2025-10-07T14:10:19.841816

oTrigger = Offset2
Dim pr As Inventor.Parameter = ThisDoc.Document.componentdefinition.parameters("Offset2")
Dim v As Double = pr.Value/2.54
pr.Expression=1
pr.Expression =Round(v,1)