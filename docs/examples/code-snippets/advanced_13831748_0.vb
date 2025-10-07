' Title: iLogic Rule to Measure Area of Two Faces
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-measure-area-of-two-faces/td-p/13831748
' Category: advanced
' Scraped: 2025-10-07T13:25:59.066228

Dim iLogicAuto = iLogicVb.Automation

Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)

Dim f As Face
f = namedEntities.FindEntity(COLD_OPENING), (HOT_OPENING)
Dim Eval As SurfaceEvaluator
Dim area As Double
Eval = f.Evaluator
area = Eval.Area
MsgBox("Face_Area : " & area)