' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T13:18:13.978278

If "Variable1" = "OptionA1" And "Variable2" = "OptionB1" OR "OptionB2" Then '(Multiple Options in the Same Line, separated by some kind of 'or' statement)
		"Component1" = True                Else If "Variable2" = "OptionB3" Then                "Component2" = True
	End If