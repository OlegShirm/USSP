Attribute VB_Name = "modWorkbookGenerator"
Option Explicit

'=========================================================
' USSP - US Stock Screener Professional
'
' Module:
'   modWorkbookGenerator
'
' Version:
'   0.1.0-dev
'
' Purpose:
'   Creates base workbook structure
'
' Milestone:
'   M2.002 VBA Core Framework
'
'=========================================================


Public Sub GenerateUSSPWorkbook()

    Dim SheetNames As Variant
    Dim i As Long
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    SheetNames = Array( _
        "Dashboard", _
        "Stock Screener", _
        "Company Data", _
        "Market Data", _
        "Watch List", _
        "Candidate List", _
        "Scan Presets", _
        "Rules Engine", _
        "Data Providers", _
        "Settings", _
        "Logs")
    
    
    'Create sheets
    For i = LBound(SheetNames) To UBound(SheetNames)
        
        If Not SheetExists(CStr(SheetNames(i))) Then
            
            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            
            ws.Name = SheetNames(i)
            
        End If
        
    Next i
    
    
    'Remove default empty sheets if necessary
    CleanDefaultSheets
    
    
    'Add project information
    With Worksheets("Dashboard")
        
        .Cells.Clear
        
        .Range("A1").Value = "USSP"
        .Range("A2").Value = "US Stock Screener Professional"
        .Range("A4").Value = "Version"
        .Range("B4").Value = "0.1.0-dev"
        
        .Range("A5").Value = "Milestone"
        .Range("B5").Value = "M2.002 VBA Core Framework"
        
        .Columns("A:B").AutoFit
        
    End With
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    
    MsgBox "USSP Workbook structure created successfully.", _
           vbInformation, _
           "USSP Generator"

End Sub



Private Function SheetExists(ByVal SheetName As String) As Boolean

    Dim ws As Worksheet
    
    SheetExists = False
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Name = SheetName Then
            SheetExists = True
            Exit Function
        End If
        
    Next ws

End Function



Private Sub CleanDefaultSheets()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Name Like "Sheet*" Then
            
            Application.DisplayAlerts = False
            ws.Delete
            
        End If
        
    Next ws

End Sub
