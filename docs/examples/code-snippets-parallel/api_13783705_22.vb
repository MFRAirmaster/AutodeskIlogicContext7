' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T14:07:55.595605

If (MiddleLetter.StartsWith("D") = True Or MiddleLetter.StartsWith("K") = True) And LastLetter.StartsWith("H") = False Then 
				Tank_Leg_Stand_TC = True	
			Else
				Tank_Leg_Stand_TC = False
			End If