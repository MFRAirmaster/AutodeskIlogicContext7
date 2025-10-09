' Title: trying to extrude a acad dxf by ilogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trying-to-extrude-a-acad-dxf-by-ilogic/td-p/13745079
' Category: advanced
' Scraped: 2025-10-09T09:06:41.279917

Sub Main()

    Dim oPartDoc As PartDocument = ThisApplication.ActiveDocument
    Dim oCompDef As PartComponentDefinition = oPartDoc.ComponentDefinition

    ' Șterge schița SketchLitere dacă există
    For Each sk As PlanarSketch In oCompDef.Sketches
        If sk.Name = "SketchLitere" Then
            sk.ExitEdit()
			sk.Delete()
            Exit For
        End If
    Next

    ' Creează schița SketchLitere pe planul XY (de obicei WorkPlanes.Item(3))
    Dim oXYPlane As WorkPlane = oCompDef.WorkPlanes.Item(3)
    Dim oSketch As PlanarSketch = oCompDef.Sketches.Add(oXYPlane)
    oSketch.Name = "SketchLitere"
    oSketch.Edit()

    ' Selectează fișierul DXF/DWG
    Dim dlg As New System.Windows.Forms.OpenFileDialog()
    dlg.Filter = "DXF or DWG files (*.dxf;*.dwg)|*.dxf;*.dwg"
    dlg.Title = "Selectează fișierul DXF/DWG"
    If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
        MessageBox.Show("Fișierul nu a fost selectat.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        oSketch.ExitEdit()
        Return
    End If

    Dim sFileName As String = dlg.FileName

    ' Lansează comanda de import DXF în schița activă (SketchLitere)
    Dim oCmdMgr As CommandManager = ThisApplication.CommandManager
    oCmdMgr.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, sFileName)
    oCmdMgr.ControlDefinitions("SketchInsertAutoCADFileCmd").Execute()

    ' Așteaptă ca utilizatorul să finalizeze importul în fereastra care s-a deschis
    MessageBox.Show("Finalizează importul în fereastra care s-a deschis, apoi apasă OK aici pentru a continua.", _
                    "Continuă", MessageBoxButtons.OK, MessageBoxIcon.Information)

    ' Închide schița pentru a putea folosi profilele
    oSketch.ExitEdit()

    ' Verifică parametrul "Adancime"
    Dim adancime As Double
    Try
        Dim oParam As UserParameter = oCompDef.Parameters.UserParameters.Item("Adancime")
        adancime = oParam.Value
    Catch ex As Exception
        MessageBox.Show("Parametrul 'Adancime' nu există. Creează-l înainte de a rula regula.", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End Try

    ' Verifică dacă există profile
    If oSketch.Profiles.Count = 0 Then
		
        MessageBox.Show("Nu s-au importat profile din DXF.", oSketch.SketchEntities.Type.ToString, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End If
    
    ' Extrudare combinată cu toate profilele
    ExtrudeAllProfiles(oCompDef, oSketch, adancime)

End Sub

Sub ExtrudeAllProfiles(oCompDef As PartComponentDefinition, oSketch As PlanarSketch, ByVal adancime As Double)
    ' Crează profil combinat cu găuri din toate profilele schiței
    Dim combinedProfile As Profile = oSketch.Profiles.AddForSolid()

    Dim extrudeDef As ExtrudeDefinition = oCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(combinedProfile, PartFeatureOperationEnum.kNewBodyOperation)
    extrudeDef.SetDistanceExtent(adancime, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)

    oCompDef.Features.ExtrudeFeatures.Add(extrudeDef)

    MessageBox.Show("Extrudare finalizată cu succes!", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information)
End Sub