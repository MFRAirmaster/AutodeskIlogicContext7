' Title: Run external rule without an open document?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/run-external-rule-without-an-open-document/td-p/10149441#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:25:42.069774

'This is the code that does the real work when your command is executed.
Sub Run_ExternalRule()
    MessageBox.Show("This command creates a new assembly from the ZeroDoc environment via External Rule.", "Create New Assembly", MessageBoxButtons.OK, MessageBoxIcon.Information)
    Run_External_iLogic_Rule_NoDoc.RunExternalRule("\\EngServer\inventor_data\Support\iLogic\CreateNewAssemblyFromZeroDoc.iLogicVb")
End Sub