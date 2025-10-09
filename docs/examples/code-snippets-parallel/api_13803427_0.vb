' Title: Trailing Zeros
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427#messageview_0
' Category: api
' Scraped: 2025-10-09T09:01:47.917202

'round values
LENGTH = Round(SheetMetal.FlatExtentsLength, 2) 
WIDTH = Round(SheetMetal.FlatExtentsWidth, 2)
'set raw part information
iProperties.Value("Project", "Stock Number") = "HRS " & iProperties.Value("Custom", "GAUGE") & " x " & WIDTH & " x " & LENGTH