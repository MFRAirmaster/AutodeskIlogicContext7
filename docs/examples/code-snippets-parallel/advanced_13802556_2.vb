' Title: Colour Coding Parts by Length in Drawings
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/colour-coding-parts-by-length-in-drawings/td-p/13802556
' Category: advanced
' Scraped: 2025-10-09T08:59:31.228687

' iLogic: Color each LEAF PART occurrence by its Part Number (drills into subassemblies)
Public Class ThisRule

  Sub Main()
    Dim doc As AssemblyDocument = ThisDoc.Document
    Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction(doc, "Unique Colors by Leaf Part Number")

    SetDesignViewRepresentation(doc)

    Dim colorByPartNr As New Dictionary(Of String, Asset)
    Dim rng As New System.Random()

    ' Iterate only real parts (leaf occurrences), not subassemblies
    For Each leaf As ComponentOccurrence In doc.ComponentDefinition.Occurrences.AllLeafOccurrences
      If leaf Is Nothing Then Continue For

      Dim pn As String = GetPartNumber(leaf)
      If String.IsNullOrWhiteSpace(pn) Then Continue For

      Dim a As Asset = Nothing
      If colorByPartNr.ContainsKey(pn) Then
        a = colorByPartNr(pn)
      Else
        a = CreateRandomColorAsset(doc, rng)
        colorByPartNr.Add(pn, a)
      End If

      ' Apply appearance to the occurrence (does not modify the source IPT)
      leaf.Appearance = a
    Next

    trans.End()
  End Sub

  Private Function CreateRandomColorAsset(doc As AssemblyDocument, rng As System.Random) As Asset
    ' Create a document-level Appearance based on "Generic"
    Dim asset As Asset = doc.Assets.Add(AssetTypeEnum.kAssetTypeAppearance, "Generic", "appearances")

    ' Random diffuse color
    Dim r As Integer = rng.Next(0, 256)
    Dim g As Integer = rng.Next(0, 256)
    Dim b As Integer = rng.Next(0, 256)

    ' Set color (handle common key names)
    Dim diffuse As ColorAssetValue = Nothing
    Try
      diffuse = TryCast(asset.Item("generic_diffuse"), ColorAssetValue)
    Catch
    End Try
    If diffuse Is Nothing Then
      Try
        diffuse = TryCast(asset.Item("DiffuseColor"), ColorAssetValue)
      Catch
      End Try
    End If
    If diffuse IsNot Nothing Then
      diffuse.Value = ThisApplication.TransientObjects.CreateColor(r, g, b)
    End If

    Return asset
  End Function

  Private Function GetPartNumber(occ As ComponentOccurrence) As String
    ' Use the referenced document’s Part Number so identical parts share a color
    Dim refDoc As Document = occ.Definition.Document
    Try
      Dim ps As PropertySet = refDoc.PropertySets.Item("Design Tracking Properties")
      Dim p As [Property] = ps.Item("Part Number")
      Return CStr(p.Value)
    Catch
      ' Fallback to occurrence name if property missing (e.g., virtual components)
      Return occ.Name
    End Try
  End Function

  Private Sub SetDesignViewRepresentation(doc As AssemblyDocument)
    Dim mgr As RepresentationsManager = doc.ComponentDefinition.RepresentationsManager
    Dim rep As DesignViewRepresentation = Nothing
    Try
      rep = mgr.DesignViewRepresentations("Unique Colors")
    Catch
      rep = mgr.DesignViewRepresentations.Add("Unique Colors")
    End Try
    rep.Activate()
  End Sub

End Class