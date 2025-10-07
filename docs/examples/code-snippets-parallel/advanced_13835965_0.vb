' Title: Ilogic select all views on a sheet and change iproperty of assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-select-all-views-on-a-sheet-and-change-iproperty-of/td-p/13835965
' Category: advanced
' Scraped: 2025-10-07T14:10:02.786239

Dim oDDoc As DrawingDocument = ThisDoc.Document
Dim oSheet As Inventor.Sheet = oDDoc.ActiveSheet
Dim sSheetNumber As String = oSheet.Name.Split(":").Last

Dim oView As DrawingView = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a drawing view of part.")
If oView Is Nothing Then Return
Dim oViewDoc As Inventor.Document = oView.ReferencedDocumentDescriptor.ReferencedDocument
Dim oViewDocCProps As Inventor.PropertySet = oViewDoc.PropertySets.Item(4)
Dim oViewDocCProp As Inventor.Property = Nothing
For Each oView In oSheet.DrawingViews
Try
	oViewDocCProp = oViewDocCProps.Item("Sheet #")
	oViewDocCProp.Value = sSheetNumber
Catch
	oViewDocCProp = oViewDocCProps.Add(sSheetNumber, "Sheet #")
End Try
Next