' Title: How to create a constraint with a part and the assembly's origin planes using the API?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-constraint-with-a-part-and-the-assembly-s-origin/td-p/13840352
' Category: advanced
' Scraped: 2025-10-07T13:10:15.584874

AssCons.AddFlushConstraint((WorkPlaneProxy)wpx1,(WorkPlane)oAsmCompDef.WorkPlanes["XY Plane"], 0);