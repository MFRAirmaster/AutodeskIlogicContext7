' Title: Check if drawing parts list is splitted
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/check-if-drawing-parts-list-is-splitted/td-p/12568593#messageview_0
' Category: api
' Scraped: 2025-10-07T13:35:29.524725

Sub main
	Dim DrawingDoc = ThisApplication.ActiveEditDocument
	Dim modelState As Integer = 0
	For Each oSheet In DrawingDoc.Sheets
		itemNumber=1
		While itemNumber < 10
			Try
				oPartsList = oSheet.PartsLists.Item(itemNumber)
				If modelState = oPartsList.ReferencedDocumentDescriptor.ReferencedModelState
					MsgBox("Parts List is split")
				Else
					MsgBox("Parts List is NOT split or this is the first occurance of it")
					modelState=oPartsList.ReferencedDocumentDescriptor.ReferencedModelState
				End If
					itemNumber = itemNumber + 1
			Catch
				Exit While
			End Try
		End While
	Next
End Sub