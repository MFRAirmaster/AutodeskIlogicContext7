' Title: ilogic rule that creating a cope of a drawing and changing its iPart instance.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-that-creating-a-cope-of-a-drawing-and-changing-its/td-p/9808596#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:09:37.418034

Sub Main()
    ' Get the active drawing document
    Dim oDrawDoc As DrawingDocument
    oDrawDoc = ThisApplication.ActiveDocument

    ' Get the referenced part document from the first view
    Dim oPartDoc As PartDocument
    Try
        Dim oView As DrawingView
        oView = oDrawDoc.ActiveSheet.DrawingViews.Item(1)
        oPartDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
    Catch
        MessageBox.Show("Error: No valid drawing view or referenced part found.", "Error")
        Exit Sub
    End Try

    ' Check if the part is an iPart factory
    If Not oPartDoc.ComponentDefinition.IsiPartFactory Then
        MessageBox.Show("The referenced document (" & oPartDoc.FullFileName & ") is not an iPart factory. Ensure the drawing references the iPart factory file, not a member file.", "Error")
        Exit Sub
    End If

    ' Get the iPart factory table
    Dim oFactory As iPartFactory
    oFactory = oPartDoc.ComponentDefinition.iPartFactory

    ' Verify that the factory has rows
    If oFactory.TableRows.Count = 0 Then
        MessageBox.Show("The iPart table is empty. Open the factory file and generate member files.", "Error")
        Exit Sub
    End If

    ' Get the path for saving drawings
    Dim oPath As String
    oPath = System.IO.Path.GetDirectoryName(oDrawDoc.FullFileName)

    ' Loop through each row in the iPart table
    Dim oRow As iPartTableRow
    For Each oRow In oFactory.TableRows
        ' Verify the member exists
        Dim oMemberFile As String
        oMemberFile = oRow.MemberPath
        If Not System.IO.File.Exists(oMemberFile) Then
            MessageBox.Show("Member file for " & oRow.MemberName & " does not exist at: " & oMemberFile & ". Regenerate member files in the factory.", "Warning")
            Continue For
        End If

        ' Activate the iPart member
        Try
            oFactory.DefaultRow = oRow
            oPartDoc.ComponentDefinition.iPart.ChangeRow(oRow.MemberName)
        Catch ex As Exception
            MessageBox.Show("Error switching to member: " & oRow.MemberName & vbCrLf & ex.Message, "Error")
            Continue For
        End Try

        ' Update the drawing
        Try
            oDrawDoc.Update
        Catch ex As Exception
            MessageBox.Show("Error updating drawing for member: " & oRow.MemberName & vbCrLf & ex.Message, "Error")
            Continue For
        End Try

        ' Generate a new file name for the drawing
        Dim oMemberName As String
        oMemberName = oRow.MemberName
        Dim oNewFileName As String
        oNewFileName = oPath & "\" & oMemberName & ".idw"

        ' Save a copy of the drawing
        Try
            oDrawDoc.SaveAs(oNewFileName, False)
        Catch ex As Exception
            MessageBox.Show("Error saving drawing for member: " & oMemberName & vbCrLf & ex.Message, "Error")
            Continue For
        End Try
    Next

    ' Revert to the first member
    Try
        oFactory.DefaultRow = oFactory.TableRows.Item(1)
        oPartDoc.ComponentDefinition.iPart.ChangeRow(oFactory.TableRows.Item(1).MemberName)
        oDrawDoc.Update
        oDrawDoc.Save
    Catch ex As Exception
        MessageBox.Show("Error reverting to first member: " & ex.Message, "Error")
    End Try

    MessageBox.Show("Drawings generated for all iPart members.", "Success")
End Sub