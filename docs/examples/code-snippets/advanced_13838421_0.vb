' Title: Can't make Plane.IntersectWithPlane work
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/can-t-make-plane-intersectwithplane-work/td-p/13838421#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:28:03.812242

Dim TG As Inv.TransientGeometry = ThisApplication.TransientGeometry
	
	Dim TargPoint As Point
	TargPoint = TG.CreatePoint(1, 1, 1)
	
	Dim StartPoint As Point
	StartPoint = TG.CreatePoint(2, 2, 1)
	
	Dim EndPoint As Point
	EndPoint = TG.CreatePoint(-1, -1, 1)

        '   Create the X-Axis vector

        Dim XYScanPlane As Inv.Plane
        XYScanPlane = TG.CreatePlaneByThreePoints(TargPoint, StartPoint, EndPoint)

        Dim X As Double
        Dim Y As Double
        Dim Z As Double
        X = StartPoint.X - TargPoint.X
        Y = StartPoint.Y - TargPoint.Y
        Z = StartPoint.Z - TargPoint.Z
        Dim XAxisVect As Inv.Vector = TG.CreateVector(X, Y, Z)
		
        Dim YZScanPlane As Inv.Plane
        YZScanPlane = TG.CreatePlane(TargPoint, XAxisVect)

        Dim YAxisLine As Inv.Line
        YAxisLine = XYScanPlane.IntersectWithPlane(YZScanPlane) ' This is likely the problem
		
		Dim RootPoint As Point = YAxisLine.RootPoint
		
		Logger.Debug("Root Point: " & RootPoint.X & ", " & RootPoint.Y & ", " & RootPoint.Z & ")")
		
		Dim YAxisUnit As Inv.UnitVector = TG.CreateUnitVector
		YAxisUnit= YAxisLine.Direction()
		Logger.Debug("Unit: " & YAxisUnit.X & ", " & YAxisUnit.Y & ", " & YAxisUnit.Z)
		
        Dim YAxisVect As Inv.Vector = TG.CreateVector(0, 0, 0)
		YAxisVect = YAxisUnit.AsVector()
		
		Logger.Debug("Vect: " & YAxisVect.X & ", " & YAxisVect.Y & ", " & YAxisVect.Z)

        Dim XZScanPlane As Inv.Plane
        XZScanPlane = TG.CreatePlane(TargPoint, YAxisVect)