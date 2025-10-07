' Title: Section &amp; Detail View Label
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080
' Category: api
' Scraped: 2025-10-07T12:48:24.301843

Dim oDoc As DrawingDocument
oDoc = ThisDoc.Document

Dim oSheet As Sheet
Dim oView As DrawingView
Dim oLabel As String
Dim RestartperSheet As Boolean

'Set this value to False if you want te rename views for all sheets
'Set this value to True if you want to rename views per sheet
RestartperSheet = True
'Set Start Label
oLabel = "A"

For Each oSheet In oDoc.Sheets
	'If True Set label Start Value per sheet
	If RestartperSheet = True
		oLabel = "A"
	End If
	For Each oView In oSheet.DrawingViews
		Logger.Info(oView.ViewType & " " & oView.Name)

		Select Case oView.ViewType
			Case kSectionDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kAuxiliaryDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kDetailDrawingViewType
				'Set view name to Label
				oView.Name = oLabel
				'Set Label to next character
				oLabel = Chr(Asc(oLabel) + 1)
			Case kProjectedDrawingViewType
				If oView.Aligned = False Then
					'Set view name to Label
					oView.Name = oLabel
					'Set Label to next character
					oLabel = Chr(Asc(oLabel) + 1)
				End If
		End Select
	Next
Next