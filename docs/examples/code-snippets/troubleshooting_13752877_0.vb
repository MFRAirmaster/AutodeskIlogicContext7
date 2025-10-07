' Title: Excel Error with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/excel-error-with-ilogic/td-p/13752877
' Category: troubleshooting
' Scraped: 2025-10-07T12:23:31.377593

Dim Filename As String = "Straight Conveyor Pricing.xlsx"
Dim SheetName As String = "Conv costing"
GoExcel.Open(Filename, SheetName)

GoExcel.CellValue("B9") = BeltWidth
GoExcel.CellValue("B10") = BeltLength
GoExcel.CellValue("B11") = Infeed_End
GoExcel.CellValue("B12") = Discharge_End
GoExcel.CellValue("B13") = Legs
GoExcel.CellValue("B14") = Feet
GoExcel.CellValue("B15") = Belting
GoExcel.CellValue("B16") = BeltingType
GoExcel.CellValue("B17") = BeltColor
GoExcel.CellValue("B18") = Flights
GoExcel.CellValue("B19") = Flight_Spacing
GoExcel.CellValue("B20") = Construction
GoExcel.CellValue("B21") = FlangeHeight
GoExcel.CellValue("B22") = Finish
GoExcel.CellValue("B24") = Drive_Style
GoExcel.CellValue("B25") = Drive_Position
GoExcel.CellValue("B26") = MotorVoltage
GoExcel.CellValue("B27") = Drive_Speed
GoExcel.CellValue("B28") = Motor
GoExcel.CellValue("B29") = Gearbox
GoExcel.CellValue("B30") = AdjustableGuides
GoExcel.CellValue("B31") = BeltLifters
GoExcel.CellValue("B32") = DripPan
GoExcel.CellValue("B33") = BeltScraper
GoExcel.CellValue("B34") = SprayBar
GoExcel.CellValue("B35") = GearboxGuards

Price = GoExcel.CellValue("B6")

GoExcel.Save
GoExcel.Close