' Title: SurfaceBody Appearance Override ignored?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/surfacebody-appearance-override-ignored/td-p/13803468#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:00:31.252185

' copy to Transient BRep
Dim transBrep As TransientBRep = ThisApplication.TransientBRep
Dim body1 As SurfaceBody =  transBrep.Copy(part1.Definition.SurfaceBodies.Item(1))
Dim body2 As SurfaceBody =  transBrep.Copy(part2.Definition.SurfaceBodies.Item(1))

' Transform the bodies (NOT THE GRAPHICS NODE!) to be in the location represented in the assembly.
Call transBrep.Transform(body1, part1.Transformation)
Call transBrep.Transform(body2, part2.Transformation)