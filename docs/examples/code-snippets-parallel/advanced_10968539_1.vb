' Title: Run AutoCad Commands from iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/run-autocad-commands-from-ilogic/td-p/10968539#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:01:21.337990

Dim acadExe = "C:\Program Files\Autodesk\AutoCAD 2020\acad.exe"
Dim dwgFile = "C:\Path\To\Drawing.dwg"
Dim scriptFile = "C:\Path\To\TestSrcipt.scr"
Dim args = dwgFile & " /b " & scriptFile
Process.Start(acadExe, args)