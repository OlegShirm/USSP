Attribute VB_Name = "modHelpers"
Option Explicit

'=========================================================
' USSP - US Stock Screener Professional
'
' Module:
'   modHelpers
'
' Version:
'   0.1.0-dev
'
' Purpose:
'   Common helper functions
'
' Milestone:
'   M2.005 Helper Functions
'
'=========================================================


'---------------------------------------------------------
' Returns True if worksheet exists
'---------------------------------------------------------
Public Function SheetExists(ByVal SheetName As String) As Boolean

    Dim ws As Worksheet

    SheetExists = False

    For Each ws In ThisWorkbook.Worksheets

        If StrComp(ws.Name, SheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If

    Next ws

End Function


'---------------------------------------------------------
' Returns True if workbook exists
'---------------------------------------------------------
Public Function WorkbookExists(ByVal WorkbookName As String) As Boolean

    Dim wb As Workbook

    WorkbookExists = False

    For Each wb In Application.Workbooks

        If StrComp(wb.Name, WorkbookName, vbTextCompare) = 0 Then
            WorkbookExists = True
            Exit Function
        End If

    Next wb

End Function


'---------------------------------------------------------
' Returns True if folder exists
'---------------------------------------------------------
Public Function FolderExists(ByVal FolderPath As String) As Boolean

    FolderExists = (Len(Dir(FolderPath, vbDirectory)) > 0)

End Function


'---------------------------------------------------------
' Returns True if file exists
'---------------------------------------------------------
Public Function FileExists(ByVal FilePath As String) As Boolean

    FileExists = (Len(Dir(FilePath)) > 0)

End Function


'---------------------------------------------------------
' Creates folder if necessary
'---------------------------------------------------------
Public Sub EnsureFolder(ByVal FolderPath As String)

    If Not FolderExists(FolderPath) Then
        MkDir FolderPath
    End If

End Sub


'---------------------------------------------------------
' Safe string conversion
'---------------------------------------------------------
Public Function NzString(ByVal Value As Variant) As String

    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = ""
    Else
        NzString = CStr(Value)
    End If

End Function


'---------------------------------------------------------
' Safe double conversion
'---------------------------------------------------------
Public Function NzDouble(ByVal Value As Variant) As Double

    Dim s As String

    If IsNull(Value) Or IsEmpty(Value) Then
        NzDouble = 0#
        Exit Function
    End If

    s = CStr(Value)

    s = Replace(s, ".", Application.DecimalSeparator)
    s = Replace(s, ",", Application.DecimalSeparator)

    If IsNumeric(s) Then
        NzDouble = CDbl(s)
    Else
        NzDouble = 0#
    End If

End Function


'---------------------------------------------------------
' Safe date conversion
'---------------------------------------------------------
Public Function NzDate(ByVal Value As Variant) As Date

    If IsDate(Value) Then
        NzDate = CDate(Value)
    Else
        NzDate = 0
    End If

End Function


'---------------------------------------------------------
' Returns current timestamp
'---------------------------------------------------------
Public Function TimeStamp() As String

    TimeStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

End Function


'---------------------------------------------------------
' Writes message to Immediate Window
'---------------------------------------------------------
Public Sub DebugLog(ByVal Message As String)

#If VBA7 Then
    Debug.Print TimeStamp() & " | " & Message
#Else
    Debug.Print Message
#End If

End Sub
