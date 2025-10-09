' Title: iLogic Toggle Folders
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-toggle-folders/td-p/13836054
' Category: advanced
' Scraped: 2025-10-09T09:05:06.578864

Dim oPane1 As BrowserPane = ThisDoc.Document.BrowserPanes("Model")

'ensure list has the right values
MultiValue.SetList("CrankType", "Crank75", "Crank100")
oCrankName = CrankType 'get value of parameter

Dim oOcc As ComponentOccurrence

'llok at each folder
For Each oFolder In oPane1.TopNode.BrowserFolders
	'get the nodes in that folder
	Dim oFolderNodes = oFolder.BrowserNode.BrowserNodes
	'evalate it folder matches parameter name and hold that value as true or false
	Dim FolderIsCranktype As Boolean = (oFolder.Name = oCrankName)
	For Each oNode As BrowserNode In oFolderNodes
		oOcc = oNode.NativeObject
		'set active status of occurrence to matchthe value of ( folder name matches parameter value)
		Component.IsActive(oOcc.name) = FolderIsCranktype
	Next
Next