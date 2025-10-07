' Title: Auto naming an excel file - Inventor 2025
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-naming-an-excel-file-inventor-2025/td-p/13779015#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:34:08.230511

Dim oDrawDoc As DrawingDocument
Dim oSheet As Sheet
Dim oDrawingView As DrawingView
Dim oDrawingViewComponent As PartDocument
Dim oDrawingViewComponentAsm As AssemblyDocument
Dim oDrawingViewComponentPres As PresentationDocument
Dim oPartDesc As GeneralNote
Dim oTextBoxes As GeneralNotes

If ThisApplication.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
	MessageBox.Show ("Please run this rule on a drawing", "Autodesk Inventor", MessageBoxButtons.OK, MessageBoxIcon.Error)
	Exit Sub
End If
	
oDrawDoc = ThisApplication.ActiveDocument

' Set a reference to the active sheet
oSheet = oDrawDoc.ActiveSheet

' Set a reference to general notes
oTextBoxes = oSheet.DrawingNotes.GeneralNotes
Try
If oSheet.DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocumentType = kPartDocumentObject Then
	oDrawingViewComponent = oSheet.DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocument  
	oSheet.Name = oDrawingViewComponent.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
Else If oSheet.DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocumentType = kAssemblyDocumentObject Then
	oDrawingViewComponentAsm = oSheet.DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocument
	oSheet.Name = oDrawingViewComponentAsm.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
Else If oSheet.DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocumentType = kPresentationDocumentObject Then
	oDrawingViewComponentPres = oSheet.DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocument
	oDrawingViewComponentAsm = oDrawingViewComponentPres.referencedfiles.item(1)
	oSheet.Name = oDrawingViewComponentAsm.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value	
End If
Catch ex As Exception 
	MessageBox.Show("Please run the rule after placing a view", "Autodesk Inventor", MessageBoxButtons.OK, MessageBoxIcon.Error)
	Exit Sub
End Try	

sText = "<StyleOverride FontSize='0.6096'><Property Document='model' PropertySet='Design Tracking Properties' Property='Part Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='5'>PART NUMBER</Property></StyleOverride><Br/><StyleOverride FontSize='0.254'><Property Document='model' PropertySet='Design Tracking Properties' Property='Description' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='29'>DESCRIPTION</Property></StyleOverride>"                
oTG = ThisApplication.TransientGeometry
oTextBox2 = oSheet.DrawingNotes.GeneralNotes.AddFitted(oTG.CreatePoint2d(0.4, 20.6), sText)