' Title: iLogic to extract Sheet Metal Data
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-extract-sheet-metal-data/td-p/13789124#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:37:34.651599

'Checks to see if kFactor, SheetStyle, FlatLength & FlatWidth parameters exist, and generates them, if not

'Source code for text parameter check

'Dim oPartDoc As PartDocument = ThisDoc.Document
'Dim oPartCompDef As PartComponentDefinition = oPartDoc.ComponentDefinition
'oParams = oPartCompDef.Parameters
'Dim oUserParams As UserParameters = oParams.UserParameters

'Try
'  p = Parameter("MyParameter")
'Catch
'  oUserParams.AddByValue("MyParameter","MyValue", UnitsTypeEnum.kTextUnits)
'End Try

Dim oApp As Inventor.Application = ThisApplication
Dim oParams As Parameters
'Dim oPartDoc As PartDocument = oApp.ActiveDocument
Dim oPartDoc As PartDocument = ThisDoc.Document
Dim oPartCompDef As PartComponentDefinition = oPartDoc.ComponentDefinition
oParams = oPartCompDef.Parameters
Dim oUserParams As UserParameters = oParams.UserParameters

'kFactor
Try
  p = Parameter("kFactor")
Catch
  oUserParams.AddByValue("kFactor", SheetMetal.ActiveKFactor, UnitsTypeEnum.kUnitlessUnits)
End Try
Parameter("kFactor") = SheetMetal.ActiveKFactor
iProperties.Value("Custom", "kFactor") = SheetMetal.ActiveKFactor

'SheetStyle
Try
  p = Parameter("SheetStyle")
Catch
  oUserParams.AddByValue("SheetStyle", SheetMetal.GetActiveStyle(), UnitsTypeEnum.kTextUnits)
End Try
Parameter("SheetStyle") = SheetMetal.GetActiveStyle() 
iProperties.Value("Custom", "Sheet Style") = SheetMetal.GetActiveStyle() 

'FlatLength
Try
  p = Parameter("FlatLength")
Catch
  oUserParams.AddByValue("FlatLength", SheetMetal.FlatExtentsLength, UnitsTypeEnum.kDefaultDisplayLengthUnits)
End Try
Parameter("FlatLength") = SheetMetal.FlatExtentsLength

'FlatWidth
Try
  p = Parameter("FlatWidth")
Catch
  oUserParams.AddByValue("FlatWidth", SheetMetal.FlatExtentsWidth, UnitsTypeEnum.kDefaultDisplayLengthUnits)
End Try
Parameter("FlatWidth") = SheetMetal.FlatExtentsWidth

'dArea
Try
  p = Parameter("dArea")
Catch
  oUserParams.AddByValue("dArea", 0, "ul")
End Try

'FlatArea
Try
  p = Parameter("FlatArea")
Catch
  oUserParams.AddByValue("FlatArea", 0, "ft^2")
End Try



'Forces update after rule runs 
iLogicVb.UpdateWhenDone = True