' Title: Stop an iLogic rule while it's running
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/stop-an-ilogic-rule-while-it-s-running/td-p/12304171#messageview_0
' Category: api
' Scraped: 2025-10-09T09:02:34.903921

For Each ...
	If IO.File.exists("C:\Temp\IF_YOU_DELETE_THIS_CODE_WILL_STOP.txt")
                MsgBox("rule " & iLogicVb.RuleName & " stopped because text file doesn't exist")                'logger.info("rule " & iLogicVb.RuleName & " stopped because text file doesn't exist")                return                 'exit sub                'goto ...
	End If
	...
Next