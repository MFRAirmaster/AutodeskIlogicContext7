' Title: SurfaceBody Appearance Override ignored?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/surfacebody-appearance-override-ignored/td-p/13803468
' Category: advanced
' Scraped: 2025-10-07T12:33:41.827885

// copy the part SrfBody to BREP and transform:
SurfaceBody bRepSrfBdy = bRep.Copy(item.SrfBody);
bRep.Transform(bRepSrfBdy, item.Transformation);

// get part appearance and copy to assembly
PartComponentDefinition partDef = (PartComponentDefinition)item.SrfBody.Parent;
PartDocument partDoc = (PartDocument)partDef.Document;
try
{
    _ = partDoc.ActiveAppearance.CopyTo(assmDoc, true);
}
catch { /* appearance already exists */ }
Asset assemblyAppearance = assmDoc.AppearanceAssets[partDoc.ActiveAppearance.Name];

// build the graphics node
GraphicsNode node = cg.AddNode(nodeID++);
SurfaceGraphics srfG = node.AddSurfaceGraphics(bRepSrfBdy);

// assign the appearance
node.Appearance = assemblyAppearance;