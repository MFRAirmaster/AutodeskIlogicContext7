' Title: what is the &quot;correct&quot; way to write this line of code for a rule?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/what-is-the-quot-correct-quot-way-to-write-this-line-of-code-for/td-p/13761867
' Category: advanced
' Scraped: 2025-10-09T09:01:32.182582

' <FireOthersImmediately>False</FireOthersImmediately>
Sub Main() ExplodePattern
'**** Define Assembly & Pattern Occurence.
Dim oDoc As AssemblyDocument
        oDoc = ThisApplication.ActiveDocument

Dim oPattern As OccurrencePattern
		oPatternInput = InputBox("Pattern Feature Name")
        oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(oPatternInput)
'**** Make each element dependent starting from the 2nd element.
Dim i As Integer

'**** Count the amount of elements after the 2nd.
    For i = 2 To oPattern.OccurrencePatternElements.Count

'**** Make all elements "Independent".
            oPattern.OccurrencePatternElements.Item(i).Independent = True
      Next

'**** Delete pattern leaving only original and independent items.
oPattern.Delete

 End Sub