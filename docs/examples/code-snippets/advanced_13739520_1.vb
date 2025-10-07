' Title: How to count specific hole features in an assembly environment
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-count-specific-hole-features-in-an-assembly-environment/td-p/13739520#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:49:01.304563

Public Class ThisRule

    Private holes As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()

    Sub Main()
        Dim oAsmDoc As AssemblyDocument = ThisDoc.Document
        Dim oAsmDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition


        For Each oOcc As ComponentOccurrence In oAsmDef.Occurrences.AllLeafOccurrences
            If Not oOcc.Suppressed Then
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim oDef As PartComponentDefinition = oOcc.Definition
                    GetHolesQty(oDef)
                End If
            End If
        Next

        Beep()

        Dim msg As String = ""
        For Each item As KeyValuePair(Of String, Integer) In holes

            Dim cell = String.Format("{0,-20}", item.Key)

            msg += String.Format("{0} : {1}{2}", cell, item.Value, System.Environment.NewLine)
        Next

        MsgBox(msg)
    End Sub


    '''Calculates holes in the part document
    '''Be careful: nested patterns are not analyzed !
    Sub GetHolesQty(ByVal oDef As PartComponentDefinition)

        Dim N As Integer = 0  'counter   
        Dim oHoles As HoleFeatures = oDef.Features.HoleFeatures
        For Each oH As HoleFeature In oHoles
            If Not oH.Suppressed Then
                AddHoleInfo(oH, 1)
            End If
        Next

        'have we any rectangular patterns ?
        Dim oRectPatterns As RectangularPatternFeatures = oDef.Features.RectangularPatternFeatures

        For Each oRPF As RectangularPatternFeature In oRectPatterns
            If oRPF.Suppressed Then Continue For

            Dim holes = oRPF.ParentFeatures.Cast(Of PartFeature).
                Where(Function(f) Not f.Suppressed And TypeOf f Is HoleFeature).ToList()

            If (holes.Count = 0) Then Continue For

            Dim elements = oRPF.PatternElements.Cast(Of FeaturePatternElement).
                Where(Function(fpe) Not fpe.Suppressed).ToList()

            For Each hole As PartFeature In holes
                AddHoleInfo(hole, elements.Count - 1)
            Next
        Next

        'have we any circular patterns ?
        Dim oCircPatterns As CircularPatternFeatures
        oCircPatterns = oDef.Features.CircularPatternFeatures

        For Each oCPF As CircularPatternFeature In oCircPatterns
            If oCPF.Suppressed Then Continue For

            Dim holes = oCPF.ParentFeatures.Cast(Of PartFeature).
                Where(Function(f) Not f.Suppressed And TypeOf f Is HoleFeature).ToList()

            If (holes.Count = 0) Then Continue For

            Dim elements = oCPF.PatternElements.Cast(Of FeaturePatternElement).
                Where(Function(fpe) Not fpe.Suppressed).ToList()

            For Each hole As PartFeature In holes
                AddHoleInfo(hole, elements.Count - 1)
            Next

        Next
    End Sub


    Private Sub AddHoleInfo(hole As HoleFeature, numberOfHoles As Integer)

        Dim name As String = "Unknown"

        If (hole.Tapped) Then
            If (TypeOf hole.TapInfo Is HoleTapInfo) Then
                Dim info As HoleTapInfo = hole.TapInfo
                name = info.ThreadDesignation
            Else
                Dim info As TaperedThreadInfo = hole.TapInfo
                name = info.ThreadDesignation
            End If

        Else
            name = String.Format("Hole: {0}", hole.HoleDiameter.Expression)
        End If

        If holes.ContainsKey(name) Then
            holes(name) += numberOfHoles
        Else
            holes.Add(name, numberOfHoles)
        End If

    End Sub
End Class