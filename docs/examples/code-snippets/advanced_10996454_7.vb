' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:33:58.654638

Dim oParalell As Boolean= curAsmOrPlane.Plane.IsParallelTo(curCompOriPlane.Plane, 0.00001)
Dim oSameDirection As Boolean= curAsmOrPlane.Plane.Normal.IsEqualTo(curCompOriPlane.Plane.Normal)

	


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