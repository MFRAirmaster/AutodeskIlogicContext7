' Title: iLogic to quickly add PDF as PNG background into current activ drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-quickly-add-pdf-as-png-background-into-current-activ/td-p/13757564#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:01:10.411531

Option Explicit

Sub Import_Excel()
    
    
    Dim xlApp As Object 'Excel.Application
    Dim xlWB As Variant
    Dim xlWS As Excel.WorkSheet
    
    Dim xlPath As String
    Dim xlFile As String
    Dim JobNo As Long
    Dim xlTotalPath As String
    
    Dim LastRow As Long
    
  
        JobNo = GetJobNumber() 'pass function into this location to make the document follow the job number dynamically.
    
        xlPath = "Path" & JobNo & "\"   
        xlFile = JobNo & "_Chart.xlsx"
        
    
    If JobNo > 37599 Then
        MsgBox "Job range exceeded, please notify programer"
        Exit Sub
    End If

    xlTotalPath = xlPath & xlFile
         'MsgBox xlTotalPath
    
         'Next section attempts to open the excel document off the given variables.
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
   
        'The below error change is required to dodge a dialog box that has no effect upon the opening of the document.
        
        On Error Resume Next
        If Dir(xlTotalPath) = "" Then
            'MsgBox "File not found " & xlTotalPath & " Make sure the nozzel fitting chart has been created and exported to job folder"
            Exit Sub
        Else
            Set xlWB = xlApp.Workbooks.Open(xlTotalPath)
        End If
        On Error GoTo 0
    
    
        If xlWB Is Nothing Then
            MsgBox "workbook not loaded"
        End If
    
            Set xlWS = xlWB.Sheets("Sheet1")
    
            LastRow = FindNotesLastRow(xlWS)
        
            ' MsgBox "LastRow: " & LastRow
            ' MsgBox "FindLast " & FindNotesLastRow(xlWS)
        
        xlWS.Range("A1:N" & LastRow).Copy
        
        Dim CtrlDef As Inventor.ControlDefinition
        Set CtrlDef = ThisApplication.CommandManager.ControlDefinitions.Item("AppPasteSpecialCmd")
        
        CtrlDef.Enabled = True
        CtrlDef.Execute '(True)

        xlWB.Application.CutCopyMode = False
        xlWB.Close (True)
        xlApp.Quit
End Sub


Function GetJobNumber()
    'Added later to parse for the current job number automatically.
    'This will then get passed into the jobNo variable within the main to grab the correct document.

    Dim odoc As DrawingDocument
    Set odoc = ThisApplication.ActiveDocument
  
    Dim iProperties As PropertySet
    Set iProperties = odoc.PropertySets.Item("Design Tracking Properties")
    
    Dim JobProperty As Property
    Set JobProperty = iProperties.Item("Project")
    
    Dim JobNumber As Long
    JobNumber = CLng(JobProperty.Value)
    
    GetJobNumber = JobNumber
    
End Function


Function FindNotesLastRow(xlWS As Excel.WorkSheet) As Long

        Dim Cell As Range
        Dim NotesRow As Long
        
        NotesRow = xlWS.Cells(xlWS.Rows.Count, "A").End(xlUp).Row
  
        
        FindNotesLastRow = NotesRow
        
End Function