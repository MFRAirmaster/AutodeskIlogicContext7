' Title: Automated Mates
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automated-mates/td-p/13756652#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:57:45.253383

Dim oParalell As Boolean= curAsmOrPlane.Plane.IsParallelTo(curCompOriPlane.Plane, 0.00001)
Dim oSameDirection As Boolean= curAsmOrPlane.Plane.Normal.IsEqualTo(curCompOriPlane.Plane.Normal)

	Dim oNV As NameValueMap


	If oParalell=True Then
'

				If oSameDirection=True Then
				oConstraint=oAsmComp.Constraints.AddFlushConstraint(curAsmOrPlane, curCompOriPlane, AbstandvonEbene(Zähler))
				End If
				If oSameDirection=False Then
				oConstraint=oAsmComp.Constraints.AddMateConstraint(curAsmOrPlane, curCompOriPlane, AbstandvonEbene(Zähler))
				End If

	  End If

	Next
	

Next