' Title: How to refresh/open the projects window VBA
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-refresh-open-the-projects-window-vba/td-p/13725927#messageview_0
' Category: api
' Scraped: 2025-10-09T09:09:38.916254

Dim oControlDef As ControlDefinition
        
Set oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppProjectsCmd")
        
Call oControlDef.Execute