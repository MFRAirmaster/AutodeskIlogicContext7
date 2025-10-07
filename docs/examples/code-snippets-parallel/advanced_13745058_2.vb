' Title: Export an Assembly using ilogic to designated excel template tabs
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/export-an-assembly-using-ilogic-to-designated-excel-template/td-p/13745058
' Category: advanced
' Scraped: 2025-10-07T14:03:26.497936

oPartsOnlyBOMView.Export(oPATH + ThisDoc.FileName(False) + " BOM" + ".xlsx", kMicrosoftExcelFormat, "PARTS ONLY")
oBOM.BOMViews.Item("Structured").Export(oPATH + ThisDoc.FileName(False) + " BOM Structured" + ".xlsx", kMicrosoftExcelFormat, "PARTS ONLY")
'==========================================================================================
'i = MessageBox.Show("Preview the BOM?", "Databar: iLogic - BOM Publisher",MessageBoxButtons.YesNo)
'If i = vbYes Then : launchviewer = 1 : Else : launchviewer = 0 : End If
'If launchviewer = 1 Then ThisDoc.Launch(oPATH + ThisDoc.FileName(False) + " BOM" + ".xls")

'excel
Dim excelApp As Excel.Application = Nothing
Dim excelWorkbook As Excel.Workbook = Nothing
Dim excelWorkbook2 As Excel.Workbook = Nothing
Dim excelWorksheet As Excel.Worksheet = Nothing
Dim excelWorksheet2 As Excel.Worksheet = Nothing

excelApp = New Excel.Application()
excelApp.Visible = False
excelWorkbook = excelApp.Workbooks.Add("C:\Temp\template.xlsx")

'parts only
excelWorkbook2 = excelApp.Workbooks.Open(oPATH + ThisDoc.FileName(False) + " BOM" + ".xlsx")
excelWorksheet2 = excelWorkbook2.Worksheets(1)
For Each oShape As Excel.Shape In excelWorksheet2.Shapes
	oShape.Select
	oShape.PlacePictureInCell
Next
excelWorksheet2.UsedRange.Copy

excelWorkbook.Worksheets("PARTS ONLY").Range("A3").PasteSpecial
excelApp.CutCopyMode = False
excelWorkbook2.Close(False)
'structured
excelWorkbook2 = excelApp.Workbooks.Open(oPATH + ThisDoc.FileName(False) + " BOM Structured" + ".xlsx")
excelWorksheet2 = excelWorkbook2.Worksheets(1)
For Each oShape As Excel.Shape In excelWorksheet2.Shapes
	oShape.Select
	oShape.PlacePictureInCell
Next
excelWorksheet2.UsedRange.Copy

excelWorkbook.Worksheets("Structured").Range("A3").PasteSpecial
excelApp.CutCopyMode = False
excelWorkbook2.Close(False)

excelApp.Visible = True
excelWorkbook.SaveAs(oPATH & oDoc.PropertySets(3)("Part Number").Value & " " & Strings.Replace(System.DateTime.Now.ToString("d"), "/", "-") & ".xlsx")