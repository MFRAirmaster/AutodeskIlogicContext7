' Title: Custom Ribbon Tabs
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-tabs/td-p/13831088
' Category: advanced
' Scraped: 2025-10-07T14:06:32.303576

Imports System.Runtime.InteropServices
Imports Inventor

' This GUID must match the ClassId and ClientId in your .addin file.
' The ComVisible(True) attribute makes this class visible to COM, which is how Inventor finds it.
<GuidAttribute("33B4722C-392E-4E34-9EA3-2E73DD0F1C13"), ComVisible(True)>
Public Class StandardAddInServer
    Implements ApplicationAddInServer

    Private m_inventorApplication As Application
    Private WithEvents m_appEvents As ApplicationEvents
    Private m_uiCreated As Boolean = False

    ' Store references to all created tabs to prevent garbage collection.
    Private m_tabs As New List(Of Inventor.RibbonTab)

    ' Use the user's local app data folder for the log file to avoid permissions issues.
    Private ReadOnly m_logFile As String = System.IO.Path.Combine(
        System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "NewTabAddin", "log.txt")

    Private Sub Log(message As String)
        ' This function will write a timestamped message to our log file.
        Try
            Dim logDirectory As String = System.IO.Path.GetDirectoryName(m_logFile)
            If Not System.IO.Directory.Exists(logDirectory) Then
                System.IO.Directory.CreateDirectory(logDirectory)
            End If

            Dim logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}{vbCrLf}"
            System.IO.File.AppendAllText(m_logFile, logMessage)
        Catch ex As Exception
            ' If logging fails, we can't do much, but this prevents a crash.
        End Try
    End Sub

    Public Sub Activate(ByVal addInSiteObject As ApplicationAddInSite, ByVal firstTime As Boolean) Implements ApplicationAddInServer.Activate
        Log("--- Add-in Activate Started ---")

        m_inventorApplication = addInSiteObject.Application
        Log("Successfully acquired Inventor Application object.")

        ' Hook into the OnReady event. This event fires once Inventor is fully loaded and idle.
        ' This is the most reliable time to create UI elements.
        m_appEvents = m_inventorApplication.ApplicationEvents
        Log("ApplicationEvents handler set up. Waiting for OnReady event.")

        Log("--- Add-in Activate Finished ---")
    End Sub

    Private Sub AppEvents_OnReady() Handles m_appEvents.OnReady
        Log("--- Inventor OnReady event fired. ---")
        ' The OnReady event can sometimes fire more than once. The flag ensures we only create the UI once.
        If Not m_uiCreated Then
            Log("UI not yet created. Proceeding to create UI.")
            CreateUserInterface()
            m_uiCreated = True
        End If
    End Sub

    Private Sub CreateUserInterface()
        Try
            Log("CreateUserInterface method started.")
            Dim uiMan As UserInterfaceManager = m_inventorApplication.UserInterfaceManager
            Log("Successfully acquired UserInterfaceManager.")

            If uiMan.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                Log("Ribbon interface is active. Proceeding to create tabs.")

                ' Define the list of ribbons where the tab should appear.
                Dim ribbonNames As String() = {"ZeroDoc", "Part", "Assembly", "Drawing"}

                For Each ribbonName In ribbonNames
                    Try
                        Dim ribbon As Ribbon = uiMan.Ribbons.Item(ribbonName)
                        Log($"Successfully acquired '{ribbonName}' ribbon.")

                        ' Use a helper function to create the tab to avoid code duplication.
                        AddTabToRibbon(ribbon)

                    Catch ex As Exception
                        Log($"!!! WARNING: Could not access or modify the '{ribbonName}' ribbon. {ex.Message}")
                    End Try
                Next
            Else
                Log("Ribbon interface is not active. Aborting UI creation.")
            End If
            Log("--- UI Creation Finished ---")
        Catch ex As Exception
            Log($"!!! FATAL ERROR during UI Creation: {ex.ToString()} !!!")
        End Try
    End Sub

    Private Sub AddTabToRibbon(ByVal targetRibbon As Ribbon)
        Try
            ' Check if the tab already exists on this specific ribbon.
            Dim existingTab As Inventor.RibbonTab = Nothing
            Try
                existingTab = targetRibbon.RibbonTabs.Item("id_Tab_OP")
            Catch
                ' Item throws an error if not found, so we catch it and proceed.
            End Try

            If existingTab IsNot Nothing Then
                Log($"Found existing 'OP' tab on ribbon '{targetRibbon.DisplayName}'. No action needed.")
            Else
                Log($"'OP' tab not found on ribbon '{targetRibbon.DisplayName}', attempting to create it now.")
                Dim newTab As Inventor.RibbonTab = targetRibbon.RibbonTabs.Add(
                    DisplayName:="OP",
                    InternalName:="id_Tab_OP",
                    ClientId:="{33B4722C-392E-4E34-9EA3-2E73DD0F1C13}")

                ' IMPORTANT: Add the new tab to our class-level list to keep it alive.
                m_tabs.Add(newTab)
                Log($"Successfully created 'OP' tab on ribbon '{targetRibbon.DisplayName}'.")
            End If
        Catch ex As Exception
            Log($"!!! ERROR adding tab to ribbon '{targetRibbon.DisplayName}': {ex.Message}")
        End Try
    End Sub

    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
        Log("--- Add-in Deactivate ---")

        ' Clean up all created tab objects.
        For Each tab As Inventor.RibbonTab In m_tabs
            If tab IsNot Nothing Then
                Try
                    tab.Delete()
                    Marshal.ReleaseComObject(tab)
                Catch ex As Exception
                    Log($"Error during tab cleanup: {ex.Message}")
                End Try
            End If
        Next
        m_tabs.Clear()

        If m_inventorApplication IsNot Nothing Then
            Marshal.ReleaseComObject(m_inventorApplication)
            m_appEvents = Nothing
            m_inventorApplication = Nothing
        End If

        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Public ReadOnly Property Automation As Object Implements ApplicationAddInServer.Automation
        Get
            Return Nothing
        End Get
    End Property

    Public Sub ExecuteCommand(ByVal commandID As Integer) Implements ApplicationAddInServer.ExecuteCommand
        ' This method is obsolete.
    End Sub
End Class