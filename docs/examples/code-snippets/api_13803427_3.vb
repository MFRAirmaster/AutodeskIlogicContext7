' Title: Trailing Zeros
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427
' Category: api
' Scraped: 2025-10-07T13:13:02.754058

'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)

'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & FormatNumber(WIDTH, 2) & " x " & FormatNumber(LENGTH, 2)