' Title: Overridden Balloons
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/overridden-balloons/td-p/13699659
' Category: troubleshooting
' Scraped: 2025-10-07T12:50:27.025750

Dim balloonType As BalloonTypeEnum
Dim balloonTypeData As Variant
Call oBalloon.GetBalloonType(balloonType, balloonTypeData)
Debug.Print "Balloon type = " & balloonType