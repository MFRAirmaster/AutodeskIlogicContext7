' Title: Run external rule without an open document?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/run-external-rule-without-an-open-document/td-p/10149441#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:10:30.958138

Module Run_External_iLogic_Rule_NoDoc

    Public Sub RunExternalRule(ByVal ExternalRuleName As String)

        Try

            ' The application object.
            Dim addIns As ApplicationAddIns = g_inventorApplication.ApplicationAddIns()

            ' Unique ID code for iLogic Addin
            Dim iLogicAddIn As ApplicationAddIn = addIns.ItemById("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}")

            ' Starts the process
            iLogicAddIn.Activate()

            ' Executes the rule
            '  iLogicAddIn.Automation.RunExternalRule(g_inventorApplication.ActiveDocument, ExternalRuleName)
            iLogicAddIn.Automation.RunExternalRule(ExternalRuleName) ' << ERROR HERE
...