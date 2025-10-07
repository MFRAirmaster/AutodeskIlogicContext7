' Title: Custom ribbon button doesn't execute VBA sub
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-button-doesn-t-execute-vba-sub/td-p/13788549#messageview_0
' Category: api
' Scraped: 2025-10-07T14:07:29.359790

Public Sub RunCleanButtonTest()
' --- Phase 1: Aggressive Cleanup ---
'Const TEST_MACRO_NAME As String = "Default.Module1.testio"
Const TEST_MACRO_NAME As String = "StampMaker.UIManager.testprompt"
Const TEST_TAB_NAME As String = "MyTestTab"
Const TEST_PANEL_NAME As String = "MyTestPanel"

' Delete the Control Definition
On Error Resume Next
ThisApplication.CommandManager.ControlDefinitions.item(TEST_MACRO_NAME).delete
Debug.Print "Old Control Definition deleted (if it existed)."
On Error GoTo 0

' Delete the Ribbon Tab
Dim zeroRibbon As Inventor.Ribbon
Set zeroRibbon = ThisApplication.UserInterfaceManager.Ribbons.item("ZeroDoc")
On Error Resume Next
zeroRibbon.RibbonTabs.item(TEST_TAB_NAME).delete
Debug.Print "Old Ribbon Tab deleted (if it existed)."
On Error GoTo 0

' --- Phase 2: Create a single, simple button ---
Debug.Print "Creating new UI..."

' Create Tab and Panel
Dim newTab As RibbonTab
Set newTab = zeroRibbon.RibbonTabs.Add("My Test", TEST_TAB_NAME, "ClientID_TestTab")
Dim newPanel As RibbonPanel
Set newPanel = newTab.RibbonPanels.Add("My Panel", TEST_PANEL_NAME, "ClientID_TestPanel")

' Create Button Definition
Dim buttonDef As ButtonDefinition
Set buttonDef = ThisApplication.CommandManager.ControlDefinitions.AddButtonDefinition( _
"Click Me", _
TEST_MACRO_NAME, _
CommandTypesEnum.kNonShapeEditCmdType)

' Add button to panel
Call newPanel.CommandControls.AddButton(buttonDef, True)

Debug.Print "--- Test UI created. Please click the 'Click Me' button on the 'My Test' tab. ---"
End Sub


Public Sub testprompt()
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
Debug.Print ("HEUREKA! it works")
End Sub