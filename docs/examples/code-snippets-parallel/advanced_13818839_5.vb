' Title: iLogic - turn on/off parts
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-turn-on-off-parts/td-p/13818839
' Category: advanced
' Scraped: 2025-10-09T08:53:22.385561

Dim oPane1 As BrowserPane = ThisDoc.Document.BrowserPanes("Model")
Dim oFolder As BrowserFolder = oPane1.TopNode.BrowserFolders.Item("MyFolder")
Dim oFolderNodes = oFolder.BrowserNode.BrowserNodes
Dim oOcc As ComponentOccurrence

For Each oNode As BrowserNode In oFolderNodes
	oOcc = oNode.NativeObject
	'turn off
	Component.IsActive(oOcc.name) = False
Next