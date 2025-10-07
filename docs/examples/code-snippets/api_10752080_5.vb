' Title: Section &amp; Detail View Label
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/section-amp-detail-view-label/td-p/10752080
' Category: api
' Scraped: 2025-10-07T12:48:24.301843

Sub Main()
    Dim oDoc As DrawingDocument
    oDoc = ThisDoc.Document

    Dim oStyleM As DrawingStylesManager
    oStyleM = oDoc.StylesManager

    Dim oStandard As DrawingStandardStyle
    oStandard = oStyleM.ActiveStandardStyle

    Dim oSheet As Sheet
    Dim oView As DrawingView
    Dim oLabel As String
    Dim RestartperSheet As Boolean

    ' Set this value to True to rename views per sheet, False for continuous labeling across all sheets
    RestartperSheet = True
    ' Set Start Label
    oLabel = "A"

    For Each oSheet In oDoc.Sheets
        ' Reset label to A for each sheet if RestartperSheet is True
        If RestartperSheet = True Then
            oLabel = "A"
        End If
        For Each oView In oSheet.DrawingViews
            Logger.Info(oView.ViewType & " " & oView.Name)
            ' Skip excluded characters for the current label
            If oStandard.ApplyExcludeCharactersToViewNames = True Then
                While oStandard.ExcludeCharacters.Contains(oLabel) = True
                    oLabel = GetNextLabel(oLabel, oStandard)
                End While
            End If
            Select Case oView.ViewType
                Case kSectionDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kAuxiliaryDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kDetailDrawingViewType
                    ' Set view name to Label
                    oView.Name = oLabel
                    ' Get next valid label
                    oLabel = GetNextLabel(oLabel, oStandard)
                Case kProjectedDrawingViewType
                    If oView.Aligned = False Then
                        ' Set view name to Label
                        oView.Name = oLabel
                        ' Get next valid label
                        oLabel = GetNextLabel(oLabel, oStandard)
                    End If
            End Select
        Next
    Next

    iLogicVb.UpdateWhenDone = True
End Sub

' Function to get the next valid label (A-Z, skipping excluded characters)
Function GetNextLabel(currentLabel As String, oStandard As DrawingStandardStyle) As String
    Dim nextLabel As String
    nextLabel = Chr(Asc(currentLabel) + 1)
    ' Ensure label stays within A-Z (ASCII 65-90)
    If Asc(nextLabel) > Asc("Z") Then
        nextLabel = "A" ' Reset to A if exceeding Z
    End If
    ' Skip excluded characters
    If oStandard.ApplyExcludeCharactersToViewNames = True Then
        While oStandard.ExcludeCharacters.Contains(nextLabel) = True
            nextLabel = Chr(Asc(nextLabel) + 1)
            If Asc(nextLabel) > Asc("Z") Then
                nextLabel = "A" ' Reset to A if exceeding Z
            End If
        End While
    End If
    GetNextLabel = nextLabel
End Function