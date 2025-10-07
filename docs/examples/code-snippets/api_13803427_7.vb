' Title: Trailing Zeros
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427
' Category: api
' Scraped: 2025-10-07T13:13:02.754058

Dim DoubleLength As Double = Round(SheetMetal.FlatExtentsLength, 2)
Dim DoubleWidth As Double = Round(SheetMetal.FlatExtentsWidth, 2)
LENGTH = DoubleLength
WIDTH = DoubleWidth

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & DoubleWidth.ToString("F2") & " x " & DoubleLength.ToString("F2")