' Title: Custom ribbon button doesn't execute VBA sub
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-button-doesn-t-execute-vba-sub/td-p/13788549#messageview_0
' Category: api
' Scraped: 2025-10-07T14:07:29.359790

Public Sub CreateMinimalMacroButton()
' --- 1. Define the names for the macro and UI elements ---
Const MACRO_FULL_NAME As String = "Module1.MyTargetSubroutine"
Const TAB_ID As String = "StampMaker_MinimalTab_1"
Const PANEL_ID As String = "StampMaker_MinimalPanel_1"

Dim uiMgr As UserInterfaceManager
Set uiMgr = ThisApplication.UserInterfaceManager

' --- 2. Clean up any old versions of these UI elements ---
' Using "On Error Resume Next" is a simple way to delete items
' without causing an error if they don't exist yet.
On Error Resume Next

' Delete the Ribbon Tab (this also removes the panel and the button on it)
uiMgr.Ribbons.item("ZeroDoc").RibbonTabs.item(TAB_ID).delete

' Delete the Control Definition (the "brain" of the button)
ThisApplication.CommandManager.ControlDefinitions.item(MACRO_FULL_NAME).delete

' Return to normal error handling
On Error GoTo 0

' --- 3. Create the new UI ---

' Get the ribbon that's visible when no document is open
Dim zeroRibbon As Ribbon
Set zeroRibbon = uiMgr.Ribbons.item("ZeroDoc")

' Create the new tab
Dim myTab As RibbonTab
Set myTab = zeroRibbon.RibbonTabs.Add("Minimal Test", TAB_ID, "ClientID_MinimalTab")

' Create a panel on the tab
Dim myPanel As RibbonPanel
Set myPanel = myTab.RibbonPanels.Add("Test Panel", PANEL_ID, "ClientID_MinimalPanel")

' Create the Macro Definition - this links the button to your code
Dim macroDef As MacroControlDefinition
Set macroDef = ThisApplication.CommandManager.ControlDefinitions.AddMacroControlDefinition(MACRO_FULL_NAME)

' Add the actual button to the panel
Call myPanel.CommandControls.AddMacro(macroDef)

' --- 4. Notify the user ---
MsgBox "Minimal macro button has been created." & vbCrLf & _
"Look for the 'Minimal Test' tab in your ribbon.", vbInformation

End Sub