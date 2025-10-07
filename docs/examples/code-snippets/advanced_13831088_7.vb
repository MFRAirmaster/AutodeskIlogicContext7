' Title: Custom Ribbon Tabs
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-tabs/td-p/13831088
' Category: advanced
' Scraped: 2025-10-07T13:23:10.519492

Private WithEvents m_buttonDef As ButtonDefinition
        Private Sub AddTabToRibbon(ByVal targetRibbon As Ribbon)

            Try
                ' Check if the tab already exists on this specific ribbon.
                Dim existingTab As Inventor.RibbonTab = Nothing
                Try
                    existingTab = targetRibbon.RibbonTabs.Item("id_Tab_OP_" & targetRibbon.InternalName)
                Catch
                    ' Item throws an error if not found, so we catch it and proceed.
                End Try

                If existingTab IsNot Nothing Then
                    Log($"Found existing 'OP' tab on ribbon '{targetRibbon.InternalName}'. No action needed.")
                Else
                    Log($"'OP' tab not found on ribbon '{targetRibbon.InternalName}', attempting to create it now.")
                    Dim newTab As Inventor.RibbonTab = Nothing
                    newTab = targetRibbon.RibbonTabs.Add(
                    DisplayName:="OP",
                    InternalName:="id_Tab_OP_" & targetRibbon.InternalName,
                    ClientId:=AddInClientID)

                    Dim opan As RibbonPanel = newTab.RibbonPanels.Add("OP", "id_Panel_OP_" & targetRibbon.InternalName, AddInClientID)
                    opan.Visible = True
                    newTab.Visible = True
                    Dim controlDefs As ControlDefinitions = m_inventorApplication.CommandManager.ControlDefinitions
                    m_buttonDef = controlDefs.AddButtonDefinition("One", "One_" & targetRibbon.InternalName, CommandTypesEnum.kNonShapeEditCmdType, AddInClientID)
                    opan.CommandControls.AddButton(m_buttonDef)
                    ' IMPORTANT: Add the new tab to our class-level list to keep it alive.
                    m_tabs.Add(newTab)
                    Log($"Successfully created 'OP' tab on ribbon '{targetRibbon.InternalName}'.")
                End If
            Catch ex As Exception
                Log($"!!! ERROR adding tab to ribbon '{targetRibbon.InternalName}': {ex.Message}")
            End Try
        End Sub