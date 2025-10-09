' Title: iLogic and Folders (again)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-and-folders-again/td-p/13843458#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:57:28.161903

For Each oFolder In oPane1.TopNode.BrowserFolders
	'get the nodes in that folder
	Dim oFolderNodes = oFolder.BrowserNode.BrowserNodes
	'evaluate if folder matches parameter name and hold that value as true or false
	Dim FolderIsCranktype As Boolean = (oFolder.Name = oCrankName)
	For Each oNode As BrowserNode In oFolderNodes
		Try
			oOcc = oNode.NativeObject
		Catch
			Logger.Info("Could not get occurrence from node")
			Continue For
		End Try
		'set active status of occurrence to match the value of (folder name matches parameter value)
		Component.IsActive(oOcc.Name) = FolderIsCranktype
	Next
Next