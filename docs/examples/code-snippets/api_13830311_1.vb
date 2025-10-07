' Title: Inventor 2026 Inventor Application, GetActiveObject is not a member of Marshal
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2026-inventor-application-getactiveobject-is-not-a/td-p/13830311
' Category: api
' Scraped: 2025-10-07T13:31:25.942270

Try
 	_invApp = Marshal.GetActiveObject("Inventor.Application")
	Catch ex As Exception
	   Try
	     Dim invAppType As Type = _
	     GetTypeFromProgID("Inventor.Application")
	 
	     _invApp = CreateInstance(invAppType)
	     _invApp.Visible = True
	
	    'Note: if you shut down the Inventor session that was started
	    'this(way) there is still an Inventor.exe running. We will use
	    'this Boolean to test whether or not the Inventor App  will
	    'need to be shut down.
	   Catch ex2 As Exception
	     MsgBox(ex2.ToString())
	     MsgBox("Unable to get or start Inventor")
	   End Try
End Try