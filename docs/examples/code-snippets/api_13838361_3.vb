' Title: ControlDefinition.Execute2 does not run unless it is the last line in the rule
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361#messageview_0
' Category: api
' Scraped: 2025-10-07T13:19:12.173680

MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)