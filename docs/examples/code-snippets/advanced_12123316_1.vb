' Title: Local iLogic Rules in Mirrored Parts?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/local-ilogic-rules-in-mirrored-parts/td-p/12123316
' Category: advanced
' Scraped: 2025-10-07T12:40:03.547860

Sub main()
	Dim oDoc As Document = ThisDoc.Document
	If Not TypeOf oDoc Is AssemblyDocument Then Exit Sub
	Dim oAsmDoc As AssemblyDocument = oDoc
	Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments
	If oRefDocs.Count = 0 Then Exit Sub
	Dim oDocs As Documents = ThisApplication.Documents
	Dim oST As Transaction = ThisApplication.TransactionManager.StartTransaction(oAsmDoc, "CopyRules")
	Dim oPBar As Inventor.ProgressBar = ThisApplication.CreateProgressBar(False, oRefDocs.Count, "Copy ryles...")
	Dim i As Integer = 1
	oPBar.Message = "Loading..."
	For Each oRefDoc As Document In oRefDocs
		oPBar.Message = "Copy ryles in document. Progress " & i & " of " & oAsmDoc.AllReferencedDocuments.Count & "."
		If TypeOf oRefDoc Is PartDocument AndAlso oRefDoc.IsModifiable Then
			Dim oPartDoc As PartDocument = oRefDoc
			Dim oDrcdComp As DerivedPartComponents = oPartDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
			If oDrcdComp.Count = 0 Then Continue For
			If Not oDrcdComp.Item(1).Definition.Mirror Then Continue For
			Dim oDrcdDoc As Document = oDrcdComp.Item(1).ReferencedDocumentDescriptor.ReferencedDocument
			Call CopyRules(oDrcdDoc, oPartDoc)
			Call CopyEventsPropSet(oDrcdDoc, oPartDoc)
			oPartDoc.Save()
		End If
		oPBar.UpdateProgress()
		i += 1
	Next
	oPBar.Close()
	oST.End()
End Sub

Private Function CopyRules(CopyFrom As Document, CopyTo As Document)
	Try
		For Each oRule As iLogicRule In iLogicVb.Automation.Rules(CopyFrom)
			Dim oCopy As iLogicRule = iLogicVb.Automation.AddRule(CopyTo, oRule.Name, "")
			oCopy.Text = oRule.Text
		Next
	Catch : End Try
End Function

Private Function CopyEventsPropSet(CopyFrom As Document, CopyTo As Document)
	Dim oPropSet As Inventor.PropertySet = CopyFrom.PropertySets("{2C540830-0723-455E-A8E2-891722EB4C3E}")
	Try
		CopyTo.PropertySets("{2C540830-0723-455E-A8E2-891722EB4C3E}").Delete()
	Catch
	End Try
	Dim newPropSet As Inventor.PropertySet = CopyTo.PropertySets.Add("_iLogicEventsRules", "{2C540830-0723-455E-A8E2-891722EB4C3E}")
	For Each prop As Inventor.Property In oPropSet
		newPropSet.Add(prop.Value, prop.Name, prop.PropId)
	Next
End Function