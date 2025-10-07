' Title: AddAngularDimension throws exception
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/addangulardimension-throws-exception/td-p/13800019
' Category: advanced
' Scraped: 2025-10-07T13:39:55.157164

var curveSegment = curve.Segments[1];
            var otherCurveSegment = otherCurve.Segments[1];

            var curveIntent = parentLayout.sheet.CreateGeometryIntent(curveSegment, IntentTypeEnum.kGeometryIntent);
            var otherCurveIntent = parentLayout.sheet.CreateGeometryIntent(otherCurveSegment, IntentTypeEnum.kGeometryIntent);