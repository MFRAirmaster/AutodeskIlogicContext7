' Title: View Identifier Modification Via iLogic - Parts List Properties
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/view-identifier-modification-via-ilogic-parts-list-properties/td-p/13758050#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:05:18.667672

'start of ilogic code
Dim oDoc As DrawingDocument:  oDoc = ThisDoc.Document

oSheets = oDoc.Sheets

For Each occ As DrawingView In oDoc.ActiveSheet.DrawingViews
	'Reference to the drawing view from the 1st selected object
	Dim oView As DrawingView = TryCast(occ, DrawingView)
	
	If oView IsNot Nothing Then	
	
		oView.ShowLabel = True
		
			'format the model iproperties	
			oPanelID = "<StyleOverride Underline='True'><Property Document='model' PropertySet='Inventor User Defined Properties' Property='Panel ID' FormatID='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' PropertyID='28'>PANEL ID</Property></StyleOverride>"
			oQty = "<StyleOverride Underline='True'> - QTY: <Property Document='model' PropertySet='Parts List Properties' Property='QTY' FormatID='{Unknown}' PropertyID='1'>QTY</Property></StyleOverride>"

			'add the custom iproperties to the view label
			oView.Label.FormattedText = oPanelID & oQty 

	End If
Next

'end of ilogic code