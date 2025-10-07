' Title: iLogic to Apply Multiple Drawing Styles Based on User Parameter (Dimensions, GD&amp;T, Leaders, Datums)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-apply-multiple-drawing-styles-based-on-user-parameter/td-p/13755516#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:30:02.896066

Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument

Dim oFCFStyle As FeatureControlFrameStyle = oDoc.StylesManager.FeatureControlFrameStyles.Item("GDT1")
Dim oDimStyle As Style = oDoc.StylesManager.Styles.Item("Dim Style1")
Dim oLStyle = oDoc.StylesManager.LeaderStyles.Item("Leader1")

For Each oSheet As Sheet In oDoc.Sheets
	For Each oFCF As FeatureControlFrame In oSheet.FeatureControlFrames
		oFCF.Style = oFCFStyle
	Next 
	For Each oDim As GeneralDimension In oSheet.DrawingDimensions.GeneralDimensions
		oDim.Style = oDimStyle
	Next 
	For Each oLeader As LeaderNote In oSheet.DrawingNotes.LeaderNotes
		oLeader.DimensionStyle = oLStyle
	Next 
Next