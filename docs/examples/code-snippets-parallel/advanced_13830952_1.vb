' Title: SetIncludeStatus gives error
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/setincludestatus-gives-error/td-p/13830952
' Category: advanced
' Scraped: 2025-10-07T14:09:44.257868

For Each sk As Sketch In pDef.Sketches
                    If Not matchStr(sk.Name, featureName) Then Continue For
                    Dim px As PlanarSketchProxy = Nothing
                    Try

                        occ.CreateGeometryProxy(sk, px)

                        view.SetIncludeStatus(px, include)
                        view.SetVisibility(px, include)

                        Return sheet.CreateGeometryIntent(px)
                    Catch ex As Exception
                        MessageBox.Show(" hata: " & ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Next