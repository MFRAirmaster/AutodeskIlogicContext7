' Title: How to create a new assembly from ZeroDoc environment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-new-assembly-from-zerodoc-environment/td-p/13746786
' Category: advanced
' Scraped: 2025-10-09T09:00:28.105689

Public Module Globals

    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

    'Unique ID for this add-in.  
    Public Const g_simpleAddInClientID As String = "d49b875a-c5e2-46af-8a59-a071d61d5b72"
    Public Const g_addInClientID As String = "{" & g_simpleAddInClientID & "}"