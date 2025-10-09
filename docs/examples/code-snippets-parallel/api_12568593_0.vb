' Title: Check if drawing parts list is splitted
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/check-if-drawing-parts-list-is-splitted/td-p/12568593#messageview_0
' Category: api
' Scraped: 2025-10-09T09:01:02.025775

Public Sub Main()
	Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
	If oDDoc Is Nothing Then Exit Sub  
	Dim oBrowserPane As BrowserPane = oDDoc.BrowserPanes.ActivePane
	Dim oTopBrowserNode As BrowserNode = oBrowserPane.TopNode   
	Dim oBrowserNode As BrowserNode    
	For Each oBrowserNode In oTopBrowserNode.BrowserNodes
		Dim oBrowserNode2 As BrowserNode
		For Each oBrowserNode2 In oBrowserNode.BrowserNodes
			Dim oNativeObject As Object = oBrowserNode2.NativeObject
			If TypeName(oNativeObject) = "PartsList" Then
				If oBrowserNode2.BrowserNodes.Count > 0 Then
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split")
				End If
				Exit For
			End If
		Next
	Next       
End Sub