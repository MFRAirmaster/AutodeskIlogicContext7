' Title: Can't make Plane.IntersectWithPlane work
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/can-t-make-plane-intersectwithplane-work/td-p/13838421
' Category: advanced
' Scraped: 2025-10-07T14:09:36.291536

Dim P As Plane
	P = TG.CreatePlaneByThreePoints(P1, P2, P3)
	
	If P1.X = P2.X And P2.X = P3.X Then
		P.Normal = TG.CreateUnitVector(1, 0, 0)
	ElseIf P1.Y = P2.Y And P2.Y = P3.Y Then
		P.Normal = TG.CreateUnitVector(0, 1, 0)
	ElseIf P1.Z = P2.Z And P2.Z = P3.Z Then
		P.Normal = TG.CreateUnitVector(0, 0, 1)
	End If