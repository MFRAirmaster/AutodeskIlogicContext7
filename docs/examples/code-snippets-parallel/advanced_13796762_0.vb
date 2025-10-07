' Title: Dynamically execute/add a Macro
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/dynamically-execute-add-a-macro/td-p/13796762#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:05:24.368568

public override void Main()
{
    //Get iLogic Automation object

    //var iLogicAuto = iLogicVb.Automation; //Strongly typed version of the line bellow
    dynamic iLogicAuto = ThisApplication.ApplicationAddIns.ItemById["{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"].Automation;

    //Create new rule
    string ruleName = "C:\\Temp\\DynamicRule.iLogicVb";
    var ruleText = @"       
Sub Main()
   MsgBox(""Dynamically created rule"")
End Sub
";
    System.IO.File.WriteAllText(ruleName, ruleText);

    //Execute rule
    var doc = ThisApplication.ActiveEditDocument;
    iLogicAuto.RunExternalRule(doc, ruleName);
}