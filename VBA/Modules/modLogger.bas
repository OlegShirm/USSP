Attribute VB_Name = "modLogger"
Option Explicit

'=========================================================
' USSP - Logger
' Module: modLogger
'=========================================================

Private Const LOG_SHEET As String = "Logs"

Public Sub InitializeLogger()

    Dim ws As Worksheet

    If Not SheetExists(LOG_SHEET) Then Exit Sub

    Set ws = Worksheets(LOG_SHEET)

    If ws.Cells(1, 1).Value <> "Date" Then

        ws.Cells.Clear

        ws.Range("A1") = "Date"
        ws.Range("B1") = "Level"
        ws.Range("C1") = "Module"
        ws.Range("D1") = "Message"

        ws.Rows(1).Font.Bold = True

    End If

End Sub

Public Sub LogInfo(ByVal ModuleName As String, _
                   ByVal Message As String)

    WriteLog "INFO", ModuleName, Message

End Sub

Public Sub LogWarning(ByVal ModuleName As String, _
                      ByVal Message As String)

    WriteLog "WARNING", ModuleName, Message

End Sub

Public Sub LogError(ByVal ModuleName As String, _
                    ByVal Message As String)

    WriteLog "ERROR", ModuleName, Message

End Sub

Private Sub WriteLog(ByVal Level As String, _
                     ByVal ModuleName As String, _
                     ByVal Message As String)

    Dim ws As Worksheet
    Dim NextRow As Long

    If Not SheetExists(LOG_SHEET) Then Exit Sub

    Set ws = Worksheets(LOG_SHEET)

    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(NextRow, 1).Value = TimeStamp()
    ws.Cells(NextRow, 2).Value = Level
    ws.Cells(NextRow, 3).Value = ModuleName
    ws.Cells(NextRow, 4).Value = Message

End Sub

