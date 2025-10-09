' Title: Combine steps in one feature
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/combine-steps-in-one-feature/td-p/13799124#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:00:32.675908

void Main()
{
    var part = ThisApplication.ActiveDocument as PartDocument;

    var compDef = part.ComponentDefinition;

    //Assumes there are at least two surface bodies in the part
    var body1 = compDef.SurfaceBodies[1];
    var body2 = compDef.SurfaceBodies[2];

    //Crate model features which should be combined in ClientFeature
    var workPoint1 = compDef.WorkPoints.AddFixed(
        ThisApplication.TransientGeometry.CreatePoint(0, 0, 0)
    );
    var workPoint2 = compDef.WorkPoints.AddFixed(
        ThisApplication.TransientGeometry.CreatePoint(0, 0, 0.5)
    );

    var workAxis1 = compDef.WorkAxes.AddByTwoPoints(workPoint1, workPoint2);

    //Keep input work points at the root of the model tree.
    workAxis1.ConsumeInputs = false;

    var workPlane1 = compDef.WorkPlanes.AddByLineAndPoint(workAxis1, workPoint2);

    var toolBodies = ThisApplication.TransientObjects.CreateObjectCollection();
    toolBodies.Add(body2);
    var combineFeature1 = compDef.Features.CombineFeatures.Add(body1, toolBodies, PartFeatureOperationEnum.kJoinOperation);

    var body3 = combineFeature1.SurfaceBody;
    var splitFeature1 = compDef.Features.SplitFeatures.SplitBody(workPlane1, body3);

    //Create ClientFeature
    var clientFeatureDefinition = compDef.Features.ClientFeatures.CreateDefinition(
        "SampleFeature",
        workPoint1,
        splitFeature1,
        null
    );
    compDef.Features.ClientFeatures.Add(clientFeatureDefinition, "YOUR-ADDIN-CLSID-GUID-HERE");
}