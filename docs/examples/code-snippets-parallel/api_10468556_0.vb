' Title: Reset view lables
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/reset-view-lables/td-p/10468556
' Category: api
' Scraped: 2025-10-07T14:15:59.996648

Sub Main()

Logger.Info(" ")
Logger.Info("RuleName: " & iLogicVb.RuleName)

'Re-numbers all view labels on all sheets
'	handles aplha or numeric
'	option to set all to aplha
'	option to set all to Numeric
'	option to set only details and sections to alpha


oScheme = "Numeric w/ details and sections as Alpha"
'oScheme = "All Numeric"
'oScheme = "All Alpha"

Dim oDoc As DrawingDocument
Dim oSheet As Sheet
Dim oView As DrawingView

'check if active document is a drawing
If ThisApplication.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
	Exit Sub
End If

oDoc = ThisApplication.ActiveDocument

i = 9
k = 1
'Rename view identifiers 
'look at all sheets
For Each oSheet In oDoc.Sheets
	'look at all views on sheet
	For Each oView In oSheet.DrawingViews

		'renumber only detail and section
		If oScheme = "All Numeric" Then
			oView.Name = i
			i = i + 1
		ElseIf oScheme = "All Alpha" Then
			oView.Name = num2Letter(i)
			i = i + 1

		Else If oScheme = "Numeric w/ details and sections as Alpha" Then

			If oView.ViewType = DrawingViewTypeEnum.kSectionDrawingViewType Or _
				oView.ViewType = DrawingViewTypeEnum.kDetailDrawingViewType Then
				oView.Name = num2Letter(i)
				
				'skip I, O, & Q
				If InStr(1, num2Letter(i), "I") > 0 Or _
					InStr(1, num2Letter(i), "O") > 0 Or _
					InStr(1, num2Letter(i), "Q") > 0 Then
					i = i + 1
					oView.Name = num2Letter(i)
				End If
				
				i = i + 1
			Else
				oView.Name = k
				k = k + 1
			End If
		End If
	Next oView
Next oSheet

End Sub

Public Function num2Letter(num As Long) As String
	'converts long to corresponding alpha
	'Ex. 1 = A, 28 = AB

	remain = num Mod 26
	whole = Fix(num / 26)

	If num < 27 Then
		If remain = 0 Then
			num2Letter = "Z"
		Else
			num2Letter = Chr(remain + 64)
		End If
	Else
		If remain = 0 Then
			num2Letter = Chr(whole + 63) & "Z"
		Else
			num2Letter = Chr(whole + 64) & Chr(remain + 64)
		End If
	End If
End Function