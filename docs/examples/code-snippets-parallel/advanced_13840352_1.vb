' Title: How to create a constraint with a part and the assembly's origin planes using the API?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-constraint-with-a-part-and-the-assembly-s-origin/td-p/13840352
' Category: advanced
' Scraped: 2025-10-09T09:01:44.885989

// Get active assembly document
AssemblyDocument asmDoc = (AssemblyDocument)Globals.InventorApp.ActiveDocument;
AssemblyComponentDefinition oAsmCompDef = asmDoc.ComponentDefinition;

// Adds a component (in this case a part) to the assembly.
ComponentOccurrence compOcc = AddComponent(PartPath);
// Get the part's definition
PartComponentDefinition compDef = (PartComponentDefinition)compOcc.Definition;

// Create proxies for the workplanes
object wpx1, wpx2;
compOcc.CreateGeometryProxy(compDef.WorkPlanes["XY Plane"], out wpx1);

// Creating the proxy for the assembly's origin XY Plane causes a COMException.
// "System.Runtime.InteropServices.COMException: 'Exception has been thrown by the target of an invocation.'"
compOcc.CreateGeometryProxy(oAsmCompDef.WorkPlanes["XY Plane"], out wpx2);

// Add the constraint to the assembly. This is never reached as an exception is thrown when creating the proxy.
AssemblyConstraints AssCons = oAsmCompDef.Constraints;
AssCons.AddFlushConstraint((WorkPlaneProxy)wpx1, (WorkPlaneProxy)wpx2, 0);