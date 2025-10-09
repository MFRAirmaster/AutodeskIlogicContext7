' Title: iLogic Rule to Measure Area of Two Faces
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-measure-area-of-two-faces/td-p/13831748#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:06:16.137540

Dim iLogicAuto = iLogicVb.Automation

Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)

Dim COLD_OPENING_Face, HOT_OPENING_Face As Face
COLD_OPENING_Face = namedEntities.FindEntity("COLD_OPENING")
HOT_OPENING_Face = namedEntities.FindEntity("HOT_OPENING")
Dim COLD_OPENING_Area, HOT_OPENING_Area, SUM_Area As Double
COLD_OPENING_Area = COLD_OPENING_Face.Evaluator.Area
HOT_OPENING_Area = HOT_OPENING_Face.Evaluator.Area
SUM_Area = COLD_OPENING_Area + HOT_OPENING_Area
MessageBox.Show("COLD_OPENING_Area= " & COLD_OPENING_Area & vbCrLf &"HOT_OPENING_Area= " & HOT_OPENING_Area & vbCrLf & "SUM_Area= " & SUM_Area, "Area")