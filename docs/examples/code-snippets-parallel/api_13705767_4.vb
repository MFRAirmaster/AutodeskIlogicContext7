' Title: iLogic Code to export Sheets to IDWS
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-export-sheets-to-idws/td-p/13705767
' Category: api
' Scraped: 2025-10-07T14:04:20.352940

Dim sSheetName As String = oSheet.Name
		If sSheetName.Contains(":") Then
			'Split Sheet name at colon character to create Array of Strings
			'then just get the first String in that Array, as the sheet's name
			'expecting only one colon character in the name, just before sheet number
			sSheetName = sSheetName.Split(":").First()
		End If
		'inserting a single space between sheet name and its Index number, for clarity
		sSheetName = sSheetName & " " & iSheetIndex.ToString()
		'combine full path with new file name and file extension, as value of new variable
		Dim sNewDrawingFFN As String = System.IO.Path.Combine(sNewPath, sSheetName & ".idw")