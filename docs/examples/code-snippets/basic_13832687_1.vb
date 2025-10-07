' Title: printing Drawings in A0,A1,A2 on our Canon Plotter
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/printing-drawings-in-a0-a1-a2-on-our-canon-plotter/td-p/13832687
' Category: basic
' Scraped: 2025-10-07T13:39:21.779665

oSheetHeight = oCurrentSheet.Height
If EqualWithinTolerance(oSheetHeight, {whatever value you want here}, {tolerance})
'use printer A 
Else
'use printer B
End If