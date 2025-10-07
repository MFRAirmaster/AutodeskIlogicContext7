' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T13:18:13.978278

If DoorType = "3F" And DoorWidth = 35.75, 33.75, 31.75, AndOr 29.75 Then                "Component1" = True        ElseIf DoorWidth = 41.75 Then                "Component2" = True        End If