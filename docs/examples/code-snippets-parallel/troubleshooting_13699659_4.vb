' Title: Overridden Balloons
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/overridden-balloons/td-p/13699659#messageview_0
' Category: troubleshooting
' Scraped: 2025-10-09T09:09:12.332068

Dim balloonType As BalloonTypeEnum
Dim balloonTypeData As Variant
Call oBalloon.GetBalloonType(balloonType, balloonTypeData)
Debug.Print "Balloon type = " & balloonType