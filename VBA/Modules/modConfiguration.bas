Attribute VB_Name = "modConfiguration"
Option Explicit

'=========================================================
' USSP - US Stock Screener Professional
'
' Module:
'   modConfiguration
'
' Purpose:
'   Configuration management
'
' Milestone:
'   M2.010
'
'=========================================================

Private Const SETTINGS_SHEET As String = "Settings"

Private Const COL_KEY As Long = 1
Private Const COL_VALUE As Long = 2


'---------------------------------------------------------
' Returns configuration value
'---------------------------------------------------------
Public Function GetConfig(ByVal Key As String, _
                          Optional ByVal DefaultValue As Variant) As Variant

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = Worksheets(SETTINGS_SHEET)

    LastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row

    For r = 2 To LastRow

        If StrComp(NzString(ws.Cells(r, COL_KEY).Value), _
                   Key, vbTextCompare) = 0 Then

            GetConfig = ws.Cells(r, COL_VALUE).Value
            Exit Function

        End If

    Next r

    GetConfig = DefaultValue

End Function


'---------------------------------------------------------
' Saves configuration value
'---------------------------------------------------------
Public Sub SetConfig(ByVal Key As String, _
                     ByVal Value As Variant)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = Worksheets(SETTINGS_SHEET)

    LastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row

    For r = 2 To LastRow

        If StrComp(NzString(ws.Cells(r, COL_KEY).Value), _
                   Key, vbTextCompare) = 0 Then

            ws.Cells(r, COL_VALUE).Value = Value
            Exit Sub

        End If

    Next r

    LastRow = LastRow + 1

    ws.Cells(LastRow, COL_KEY).Value = Key
    ws.Cells(LastRow, COL_VALUE).Value = Value

End Sub


'---------------------------------------------------------
' Creates Settings sheet headers
'---------------------------------------------------------
Public Sub InitializeConfiguration()

    Dim ws As Worksheet

    Set ws = Worksheets(SETTINGS_SHEET)

    If ws.Cells(1, 1).Value = "" Then

        ws.Cells(1, 1).Value = "Key"
        ws.Cells(1, 2).Value = "Value"

        ws.Columns("A:B").AutoFit

    End If

End Sub

