' Title: iLogic vs Sheet Metal Thickness: Different results between Inventor 2024.2 and 2026.1
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-vs-sheet-metal-thickness-different-results-between/td-p/13830817
' Category: advanced
' Scraped: 2025-10-09T09:09:08.971172

Private Sub GetFlatPatternDimensions()

    Dim oSheetMetalCompDef As SheetMetalComponentDefinition = ThisDoc.Document.ComponentDefinition
    Dim UOM As UnitsOfMeasure = ThisDoc.Document.UnitsOfMeasure
    Dim oFlatPattern As FlatPattern = oSheetMetalCompDef.FlatPattern

    ' Sheet Metal Thickness value is just converted to 'mm' but not rounded.
    Dim dThickness As Double
    Try
        dThickness = oSheetMetalCompDef.Thickness.Value
        dThickness = UOM.ConvertUnits(dThickness, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits)
    Catch
        Logger.Warn(iLogicVb.RuleName & " | " & ThisDoc.FileName(True) & " | Thickness parameter not found.")
        dThickness = 0
    End Try

    ' Length and Width are of data type 'Double' which may contain precision errors way beyond the decimal seperator.
    ' CompanyX uses .1 precision on drawings and .3 decimal precision when dimensioning parts in 3D therefore first step is
    ' to convert values to 'mm', round values to 1 decimals to match drawing precision and later on use 'Ceil()'-Method
    ' to round up to next full integer.
    Dim dFlatPatternLength As Double
    Try
        dFlatPatternLength = Round(UOM.ConvertUnits(oFlatPattern.Length, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits), 1)
        dFlatPatternLength = Ceil(dFlatPatternLength)
    Catch
        Logger.Warn(iLogicVb.RuleName & " | " & ThisDoc.FileName(True) & " | Flat pattern length not found.")
        dFlatPatternLength = 0
    End Try

    Dim dFlatPatternWidth As Double
    Try
        dFlatPatternWidth = Round(UOM.ConvertUnits(oFlatPattern.Width, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits), 1)
        dFlatPatternWidth = Ceil(dFlatPatternWidth)
    Catch
        Logger.Warn(iLogicVb.RuleName & " | " & ThisDoc.FileName(True) & " | Flat pattern width not found.")
        dFlatPatternWidth = 0
    End Try

    Dim sSemiFinishedProductDimension As String

    ' Make sure that higher dimension value (Lenght or Width) comes first after thickness value
    If dFlatPatternLength > dFlatPatternWidth Then

        sSemiFinishedProductDimension = "Blech " & _
                                        dThickness & _
                                        " x " & _
                                        dFlatPatternLength & _
                                        " x " & _
                                        dFlatPatternWidth

    Else

        sSemiFinishedProductDimension = "Blech " & _
                                        dThickness & _
                                        " x " & _
                                        dFlatPatternWidth & _
                                        " x " & _
                                        dFlatPatternLength

    End If

    ' Write dimension to iProperty
    iProperties.Value("Custom", "Semifinished product_dimension") = sSemiFinishedProductDimension
    Logger.Info(iLogicVb.RuleName & " | " & ThisDoc.FileName(True) & " | Semifinished Dimension: " & sSemiFinishedProductDimension)
End Sub