' Title: iLogic Remove all Event Triggers &amp; add new Event Trigger
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-remove-all-event-triggers-amp-add-new-event-trigger/td-p/13234890#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:06:50.006076

If ThisDoc.Document.DocumentInterests.HasInterest("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}") = True
	Logger.Debug("hasInterest")
ElseIf ThisDoc.Document.DocumentInterests.HasInterest("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}") = False
	Logger.Debug("hasInterest False ")
	iLogicVb.Automation.AddRule(ThisDoc.Document, "TriggerRule", "'This rule is used to activate the iLogic environment.")
	Logger.Info("The local rule TriggerRule was created to activate the iLogic environment.")
End If