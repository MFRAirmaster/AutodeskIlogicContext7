' Title: CUSTOM VIEW ORIENTATION ALTERNATIVE
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-view-orientation-alternative/td-p/13393159
' Category: advanced
' Scraped: 2025-10-07T14:07:54.930591

With oPosVectorX
	Dim dCoors() As Double = {-.X, -.Y, -.Z }
	.PutUnitVectorData(dCoors)	
End With
oCamera.UpVector = oPosVectorX