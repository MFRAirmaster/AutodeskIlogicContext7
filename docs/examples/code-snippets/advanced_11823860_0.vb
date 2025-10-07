' Title: Access IManagedDrawingView from External Rule
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/access-imanageddrawingview-from-external-rule/td-p/11823860
' Category: advanced
' Scraped: 2025-10-07T13:08:17.843738

oDoc = ThisDoc.Document
SOP = StandardObjectFactory.Create(oDoc)
oManagedDView = SOP.ThisDrawing.Sheet("SheetName").View("ViewName")