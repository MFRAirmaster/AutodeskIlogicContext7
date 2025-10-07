' Title: Export an Assembly using ilogic to designated excel template tabs
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-an-assembly-using-ilogic-to-designated-excel-template/td-p/13745058
' Category: advanced
' Scraped: 2025-10-07T14:03:26.497936

'BOM Publisher
oDoc = ThisDoc.ModelDocument
If oDoc.DocumentType = kPartDocumentObject Then
MessageBox.Show("You need to be in an Assembly to Export a BOM", "Databar: iLogic - BOM Publisher")
Return
End If
oDoc = ThisApplication.ActiveDocument
Dim oBOM As BOM
oBOM = oDoc.ComponentDefinition.BOM
'Options.Value("Author") = iProperties.Value("Summary", "Author")
'==========================================================================================
'You can change the output path by editing oPATH below
oPATH = ("c:\temp\") 'If you change this, remember to keep a \ at the end
'==========================================================================================
'STRUCTURED BoM ===========================================================================
' the structured view to 'all levels'
'oBOM.StructuredViewFirstLevelOnly = False

'
'==========================================================================================
'PARTS ONLY BoM ===========================================================================
' Make sure that the parts only view Is enabled.
oBOM.PartsOnlyViewEnabled = True
Dim oPartsOnlyBOMView As BOMView
oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")
' Export the BOM view to an Excel file
'oPartsOnlyBOMView.Export (oPATH + "BOM3.xls", kMicrosoftExcelFormat)
oPartsOnlyBOMView.Export (oPATH + ThisDoc.FileName(False) + " BOM" + ".xls", kMicrosoftExcelFormat, "PARTS ONLY")
'==========================================================================================
i = MessageBox.Show("Preview the BOM?", "Databar: iLogic - BOM Publisher",MessageBoxButtons.YesNo)
If i = vbYes Then : launchviewer = 1 : Else : launchviewer = 0 : End If
If launchviewer = 1 Then ThisDoc.Launch(oPATH + ThisDoc.FileName(False) + " BOM" + ".xls")