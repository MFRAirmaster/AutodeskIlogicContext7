' Title: Code to measure length and width
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/code-to-measure-length-and-width/td-p/13792602
' Category: api
' Scraped: 2025-10-09T09:02:26.048610

Dim oInvApp As Inventor.Application = ThisApplication
Dim oCM As CommandManager = oInvApp.CommandManager
Dim oUOM As UnitsOfMeasure = oInvApp.ActiveDocument.UnitsOfMeasure
'Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kFootLengthUnits
Dim eLeng As UnitsTypeEnum = UnitsTypeEnum.kCentimeterLengthUnits

Do
	Dim oFace As Object
	oFace = oCM.Pick(SelectionFilterEnum.kPartFaceFilter, "Select Face...")
	If oFace Is Nothing Then Exit Do
	Dim oBox As Box2d = oFace.Evaluator.ParamRangeRect
	Dim oSizes(1) As Double
	oSizes(0) = Round(oUOM.ConvertUnits(oBox.MaxPoint.X - oBox.MinPoint.X,
										eLeng, oUOM.LengthUnits), 3)
	oSizes(1) = Round(oUOM.ConvertUnits(oBox.MaxPoint.Y - oBox.MinPoint.Y,
										eLeng, oUOM.LengthUnits), 3)
	MessageBox.Show("Length - " & Abs(oSizes.Max) & vbLf &
					"Width - " & Abs(oSizes.Min), "Surface size:")
	Logger.Info("Length - " & Abs(oSizes.Max))
	Logger.Info("Width - " & Abs(oSizes.Min))
Loop