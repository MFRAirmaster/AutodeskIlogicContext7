' Title: automate the creation of individual part drawings
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automate-the-creation-of-individual-part-drawings/td-p/13755659#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:13:44.531663

Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument

Dim fileTemplate As String = "C:\Path\To\Template.idw"

For Each oLeaf As ComponentOccurrence In oDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
	
	
	Dim oDDoc1 As DrawingDocument = ThisApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, fileTemplate, True)
	oDDoc1.Activate()
	
	Dim oSheet As Sheet = oDDoc1.Sheets.Item(1)
	
	Dim xPos As Double = oSheet.Width / 2 
	Dim yPos As Double = oSheet.Height / 2
	
	Dim oPartDoc As PartDocument = oLeaf.Definition.Document
	Dim oPos As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(xPos, yPos) 'adjust as needed
	Dim oScale As Double = 1 'full scale
	Dim oViewOrient As ViewOrientationTypeEnum = ViewOrientationTypeEnum.kFrontViewOrientation
	Dim oViewStyle As DrawingViewStyleEnum = DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle 'change as necessary
	
	Dim oView As DrawingView = oSheet.DrawingViews.AddBaseView(oPartDoc, oPos, oScale, oViewOrient, oViewStyle)
	
	Dim oTopViewPos As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(xPos, yPos + oView.Height*1.5) 'adjust as needed
	
	Dim oTopView As DrawingView = oSheet.DrawingViews.AddProjectedView(oView, oTopViewPos, oViewStyle)
	
	'auto-dimension
	
	oSheet.RetrieveAnnotations(oView)
	oSheet.RetrieveAnnotations(oTopView)
	
	'save as .idw, .dwg, .pdf
	
	Dim fileName As String = oPartDoc.DisplayName
	
	oDDoc1.SaveAs("C:\Path" & fileName & ".idw", False)
	
	Dim oFilePath As String = ThisDoc.Path
	Dim oFileName As String = ThisDoc.FileName
	
	Dim oSavePath As String = oFilePath & "\" & oFileName
	
	oDDoc1.SaveAs(oSavePath & ".pdf", True)
	oDDoc1.SaveAs(oSavePath & ".dwg", True)
	
Next