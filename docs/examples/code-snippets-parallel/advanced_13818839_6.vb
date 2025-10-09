' Title: iLogic - turn on/off parts
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-turn-on-off-parts/td-p/13818839
' Category: advanced
' Scraped: 2025-10-09T08:53:22.385561

Dim oDoc As Inventor.Document = ThisApplication.ActiveDocument
Dim oPane As BrowserPane = oDoc.BrowserPanes.Item("Model")
Dim oFolder As BrowserFolder = oPane.TopNode.BrowserFolders.Item("Group1")
oDoc.SelectSet.Clear()
oFolder.BrowserNode.DoSelect()
ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyCompSuppressionCtxCmd").Execute()
oDoc.SelectSet.Clear()