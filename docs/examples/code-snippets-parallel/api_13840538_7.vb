' Title: Looking for some &quot;fun&quot; iLogic code for a T-Shirt
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/looking-for-some-quot-fun-quot-ilogic-code-for-a-t-shirt/td-p/13840538#messageview_0
' Category: api
' Scraped: 2025-10-09T09:01:44.887438

Sub EngineeringProcess()
    ' Engineering is spelled: C-H-A-N-G-E
    
    Design = CreateInitialConcept()
    SendToClient()
    
    While (ProjectNotCancelled)
        changeRequest = Client.SendEmail()
        
        Select Case changeRequest.Severity
            Case "Minor tweak"
                RedoEverything()
            Case "Small adjustment"
                StartFromScratch()
            Case "Just one little thing"
                QuestionCareerPath()
        End Select
        
        UpdateRevisionNumber() ' Now at Rev ZZ.47
        
        If (design.LooksLikeOriginalConcept = False) Then
            Console.WriteLine("Nailed it! ��")
        End If
        
        WaitForNextChangeRequest(milliseconds:=3)
    End While
    
    ' Note: ProjectNotCancelled always returns True
End Sub