' Title: How to refresh/open the projects window VBA
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-refresh-open-the-projects-window-vba/td-p/13725927
' Category: api
' Scraped: 2025-10-07T13:25:33.446484

Dim oControlDef As ControlDefinition
        
Set oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppProjectsCmd")
        
Call oControlDef.Execute