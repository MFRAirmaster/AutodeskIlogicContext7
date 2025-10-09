' Title: Modify Bom Table Settings (add Custom iProperty Column)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/modify-bom-table-settings-add-custom-iproperty-column/td-p/13792951#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:54:28.742019

If Not TypeOf ThisDoc.Document Is AssemblyDocument Then MsgBox("Not Assy!") : Return

Dim aDoc As AssemblyDocument = ThisDoc.Document
Dim aDef As AssemblyComponentDefinition = aDoc.ComponentDefinition
Dim aBom As BOM = aDef.BOM
Logger.Info("aBom: " & If(aBom Is Nothing, "-", "+"))
Logger.Info("BOMViews Count: " & aBom.BOMViews.Count)

Dim aBomCstSetFilePath As String = "C:\Temp\BomCstSets.xml" ' aDoc.FullFilename & "_BomCstSets.xml" ' System.IO.Path.ChangeExtension(aDoc.FullFilename, "xml")
Logger.Info("aBomCstSetFilePath: " & aBomCstSetFilePath)

aBom.ExportBOMCustomization(aBomCstSetFilePath) ' Unspecified error (Exception from HRESULT: 0x80004005 (E_FAIL))