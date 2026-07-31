Attribute VB_Name = "modUIManager"

Option Explicit

'====================================================
' USSP UI Manager
' Version: 0.1.0-dev
'====================================================

Public Sub ApplyUSSPStyle()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        With ws.Cells
            .Font.Name = "Calibri"
            .Font.Size = 11
        End With
        
        ws.Cells.EntireColumn.AutoFit
        
    Next ws
    
    MsgBox "USSP UI Style applied successfully.", _
           vbInformation, _
           "USSP UI Manager"

End Sub
