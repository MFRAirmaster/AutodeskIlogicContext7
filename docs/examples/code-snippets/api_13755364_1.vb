' Title: Error 0x800AC472
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/error-0x800ac472/td-p/13755364
' Category: api
' Scraped: 2025-10-07T13:14:01.828602

Dim oExcelPath As Object = "C:\Users\Lyle\SCR Solutions\Shared OneDrive - Draftsmen Access\Configurator files\Straight Conveyor Configurator\CAD\Straight Conveyor Pricing Test.xlsx"
'	Dim oExcelPath As String = "Straight Conveyor Pricing.xlsx"
	Dim oExcel As Object = CreateObject("Excel.Application")
	oExcel.Visible = False
	oExcel.DisplayAlerts = False
	Dim oWB As Object = oExcel.Workbooks.Open(oExcelPath)
	Dim oWS As Object = oWB.Sheets(1)
	oWS.Activate


oWS.Cells(9, 2) = BeltWidth
oWS.Cells(10, 2).Value = BeltLength
oWS.Cells(11, 2).Value = Infeed_End
oWS.Cells(12, 2).Value = Discharge_End
oWS.Cells(13, 2).Value = Legs
oWS.Cells(14, 2).Value = Feet
oWS.Cells(15, 2).Value = Belting
'oWS.Cells(16, 2).Value = BeltingType
'oWS.Cells(17, 2).Value = BeltColor
'oWS.Cells(18, 2).Value = Flights
'oWS.Cells(19, 2).Value = Flight_Spacing
oWS.Cells(20, 2).Value = Construction
'oWS.Cells(21, 2).Value = FlangeHeight
oWS.Cells(22, 2).Value = Finish
oWS.Cells(24, 2).Value = Drive_Style
'oWS.Cells(25, 2).Value = Drive_Position
'oWS.Cells(26, 2).Value = MotorVoltage
'oWS.Cells(27, 2).Value = Drive_Speed
oWS.Cells(28, 2).Value = Motor
oWS.Cells(29, 2).Value = Gearbox
oWS.Cells(30, 2).Value = AdjustableGuides
oWS.Cells(31, 2).Value = BeltLifters
oWS.Cells(32, 2).Value = DripPan
'oWS.Cells(33, 2).Value = BeltScraper
oWS.Cells(34, 2).Value = SprayBar
oWS.Cells(35, 2).Value = GearboxGuards

Price = oWS.Cells(6, 2).Value
	oWB.Close (True)
	oExcel.Quit
	oExcel = Nothing