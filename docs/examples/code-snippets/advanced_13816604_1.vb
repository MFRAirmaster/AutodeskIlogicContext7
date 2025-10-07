' Title: VB API to add custom IProperties to BOM
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/vb-api-to-add-custom-iproperties-to-bom/td-p/13816604#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:05:31.586828

Private Sub btnAddCustomPropsToBOM_Click(sender As Object, e As EventArgs)
      Try
          Dim oDoc As AssemblyDocument = TryCast(_inventorApp.ActiveDocument, AssemblyDocument)
          If oDoc Is Nothing Then
              MessageBox.Show("Active document is not an assembly.", "Error")
              Return
          End If

          Dim oBOM As BOM = oDoc.ComponentDefinition.BOM
          oBOM.PartsOnlyViewEnabled = True

          ' Ensure the "Parts Only" BOM view exists
          Dim oBOMView As BOMView = Nothing
          Try
              oBOMView = oBOM.BOMViews.Item("Parts Only")
          Catch
              MessageBox.Show("The 'Parts Only' BOM view is not available.", "Error")
              Return
          End Try

          ' List of standard custom iProperties to always add as BOM columns
          Dim standardProps As String() = {
              "Ancho", "Tipo_Parte", "Proveedor", "Nombre_Parte", "CostoLaser", "TestGithub"
          }

          Dim addedCount As Integer = 0
          Dim alreadyInBOM As New List(Of String)
          For Each propName In standardProps
              Try
                  oBOMView.BOMColumns.Add(2, propName) ' 2 = User Defined Property
                  addedCount += 1
              Catch ex As Exception
                  alreadyInBOM.Add(propName)
              End Try
          Next

          Dim msg As String = $"Added {addedCount} custom iProperties to the BOM."
          If alreadyInBOM.Count > 0 Then
              msg &= vbCrLf & "Already present or could not be added: " & String.Join(", ", alreadyInBOM)
          End If
          MessageBox.Show(msg, "Result")
      Catch ex As Exception
          MessageBox.Show("Error adding custom iProperties to BOM: " & ex.Message)
      End Try
  End Sub