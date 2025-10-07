' Title: SetIncludeStatus gives error
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/setincludestatus-gives-error/td-p/13830952#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:17:14.729161

For Each pl In pDef.WorkPlanes
                    If Not matchStr(pl.Name, featureName) Then Continue For
                    Dim px As WorkPlaneProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(pl, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch ex As Exception
                        MessageBox.Show(" hata: " & ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Next

                For Each ax In pDef.WorkAxes
                    If Not matchStr(ax.Name, featureName) Then Continue For
                    Dim px As WorkAxisProxy = Nothing
                    Try
                        occ.CreateGeometryProxy(ax, px)
                        view.SetIncludeStatus(px, include)
                        Return sheet.CreateGeometryIntent(px)
                    Catch ex As Exception
                        MessageBox.Show(" hata: " & ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Next