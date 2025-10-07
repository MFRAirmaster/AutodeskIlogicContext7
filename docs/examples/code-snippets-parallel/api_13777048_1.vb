' Title: VBA : Using GetObject in Inventor 2024
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/vba-using-getobject-in-inventor-2024/td-p/13777048#messageview_0
' Category: api
' Scraped: 2025-10-07T14:11:41.607188

Option Explicit
Sub OpenExcel()
    Dim oExcelApp As Object
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible = True
    
    Dim oWB As Object
    Set oWB = oExcelApp.Workbooks.Add()
    
    Dim oWS As Object
    Set oWS = oWB.Worksheets.Add()
    oWS.Name = "My New Worksheet"
    
    Call MsgBox("Review New Instance Of Excel, And New, Renamed Sheet In It.", , "")
    
    Call oWB.Close
    Call oExcelApp.Quit
    
    Set oWS = Nothing
    Set oWB = Nothing
    Set oExcelApp = Nothing
End Sub