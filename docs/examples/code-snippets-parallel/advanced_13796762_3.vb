' Title: Dynamically execute/add a Macro
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/dynamically-execute-add-a-macro/td-p/13796762
' Category: advanced
' Scraped: 2025-10-09T09:03:30.930021

public override void Main()
{
    // Get iLogic Automation Object
    var iLogicAuto = ThisApplication.ApplicationAddIns.ItemById["{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"].Automation;

    // Create new rule
    string ruleName = "C:\\Temp\\DynamicRule.iLogicVb";
    string ruleText = $"Sub Main()\nMsgBox(\"Dynamically created rule\")\nEnd Sub";

    System.IO.File.WriteAllText(ruleName, ruleText);
   
    var doc = ThisApplication.ActiveEditDocument;

    // Execute rule
    iLogicAuto.GetType().InvokeMember("RunExternalRule", BindingFlags.Public | BindingFlags.InvokeMethod, null, iLogicAuto, [doc, ruleName]);
}