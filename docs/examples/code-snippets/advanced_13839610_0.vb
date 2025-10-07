' Title: 3D sketch line between two UCS center points
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/3d-sketch-line-between-two-ucs-center-points/td-p/13839610#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:29:37.139546

Dim doc As PartDocument = ThisDoc.Document
Dim def As PartComponentDefinition = doc.ComponentDefinition

Dim wpA1 = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, "Select work point A1")
Dim wpB1 = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, "Select work point B1")


Dim sketch3D As Sketch3D = def.Sketches3D.Add()
sketch3D.SketchLines3D.AddByTwoPoints(wpA1, wpB1)
sketch3D.Name = ("Strut 01")