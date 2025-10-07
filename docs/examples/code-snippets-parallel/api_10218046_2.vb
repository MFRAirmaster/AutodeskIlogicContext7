' Title: iLogic - Check active project is true
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-check-active-project-is-true/td-p/10218046#messageview_0
' Category: api
' Scraped: 2025-10-07T14:11:29.561532

Sub Main
	Const sProjPath As String = "C:\Temp\"
	Const sProjFileName As String = "MyInventorProject.ipj"
	Const sProjFile As String = "C:\Temp\MyInventorProject.ipj"
	Dim oProjMgr As Inventor.DesignProjectManager = ThisApplication.DesignProjectManager
	'get FullFileName of active DesignProject
	Dim sActiveProj As String = oProjMgr.ActiveDesignProject.FullFileName
	'only react if it is not the one expected / wanted
	If Not sActiveProj = sProjFile Then
		'Log it, or let user know
		Logger.Info("Active Project File Is:  " & sActiveProj)
		'MessageBox.Show("Active Project File Is:  " & sActiveProj)
		'try to find / get the DesignProject we want, by its FullFileName
		Dim oMyProj As Inventor.DesignProject = Nothing
		Try
			oMyProj = oProjMgr.DesignProjects.ItemByName(sProjFile)
		Catch
			Logger.Warn("Specified Inventor DesignProject Not Fount In Projects Collection!")
		End Try
		'if it was not found in that collection, then add it
		If oMyProj Is Nothing Then
			Try
				oMyProj = oProjMgr.DesignProjects.AddExisting(sProjFile)
			Catch
				Logger.Error("Error Adding Specified DesignProject To Projects Collection!")
			End Try
		End If
		'if we now have it, then activate it
		If oMyProj IsNot Nothing Then
			Try
				'True = Set as Default Project
				oMyProj.Activate(True)
			Catch
				Logger.Error("Error Activating Specified DesignProject!")
			End Try
		End If
	End If
End Sub