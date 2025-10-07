' Title: CreateGeometryIntent creates kNoPointIntent type
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/creategeometryintent-creates-knopointintent-type/td-p/13835878
' Category: advanced
' Scraped: 2025-10-07T12:40:38.462572

Dim v1_CenterLine2_GI As GeometryIntent = IncludeWorkFeatureAndGetIntent(v1,
                                                                                     partName:="aaa",
                                                                                     featureName:="X Axis",
                                                                                     include:=True,
                                                                                     containsMatch:=True,
                                                                                     searchAxes:=True)