' Title: iLogic - Check active project is true
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-check-active-project-is-true/td-p/10218046
' Category: api
' Scraped: 2025-10-07T13:10:07.323877

Public Sub InventorProjectFile(oInvapp As Inventor.Application)

        Dim proj As DesignProject

        Try
            proj = oInvapp.DesignProjectManager.DesignProjects.ItemByName("C:\#######.ipj")
            proj.Activate()
        Catch ex As Exception
            If Not oInvapp.DesignProjectManager.ActiveDesignProject.FullFileName = "C:\#######.ipj" Then
                MsgBox("The wrong Project file has been Loaded !" & vbCrLf & "The Current Project file = " & oInvapp.DesignProjectManager.ActiveDesignProject.Name & vbCrLf & "Load the " & "C:\#######.ipj" & " file", vbCritical)
            End If
        End Try
end sub