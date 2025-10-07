' Title: AddAngularDimension throws exception
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/addangulardimension-throws-exception/td-p/13800019
' Category: advanced
' Scraped: 2025-10-07T14:08:53.779455

// 1. Find the shared point between the two curves
Point2d sharedPoint = FindSharedEndpoint(curve1, curve2);
if (sharedPoint == null)
    throw new Exception("Drawing curves do not share a common endpoint.");

// 2. Determine the other endpoints
Point2d otherPoint1 = GetOtherEndpoint(curve1, sharedPoint);
Point2d otherPoint2 = GetOtherEndpoint(curve2, sharedPoint);

// 3. Create GeometryIntents for the three points
GeometryIntent intentShared = parentLayout.sheet.CreateGeometryIntent(sharedPoint, IntentTypeEnum.kPointIntent);
GeometryIntent intent1 = parentLayout.sheet.CreateGeometryIntent(otherPoint1);
GeometryIntent intent2 = parentLayout.sheet.CreateGeometryIntent(otherPoint2);

var tg = parentLayout.m_application.TransientGeometry;
var newPos = tg.CreatePoint2d(intentShared.PointOnSheet.X - 1, intentShared.PointOnSheet.Y - 2);

//DrawingSketch tempSketch = parentLayout.drawingView.Sketches.Add();
DrawingSketch tempSketch = parentLayout.sheet.Sketches.Add();
tempSketch.Edit();
SketchLine tempLine = tempSketch.SketchLines.AddByTwoPoints(sharedPoint, newPos);
SketchLine curveLine = tempSketch.SketchLines.AddByTwoPoints(sharedPoint, otherPoint1);
SketchLine otherLine = tempSketch.SketchLines.AddByTwoPoints(sharedPoint, otherPoint2);
tempSketch.SketchPoints.Add(newPos);

tempSketch.ExitEdit();

// 4. Add angular dimension using 3-point method
//Point2d position = /* your text position here */;
AngularGeneralDimension angDim = genDims.AddAngular(
    newPos,
    intentShared,
    intent1,
    intent2);