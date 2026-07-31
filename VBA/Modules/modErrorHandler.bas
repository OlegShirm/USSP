Attribute VB_Name = "modErrorHandler"
Option Explicit

'=========================================================
' USSP
'
' Module:
'   modErrorHandler
'
' Purpose:
'   Centralized error handling
'
'=========================================================

Public Sub HandleError(ByVal ProcedureName As String, _
                       ByVal Ex As ErrObject)

    Dim Msg As String

    Msg = "Procedure: " & ProcedureName & vbCrLf & _
          "Error " & Ex.Number & vbCrLf & _
          Ex.Description

    LogError ProcedureName, _
             "Error " & Ex.Number & ": " & Ex.Description

    MsgBox Msg, _
           vbCritical, _
           APP_NAME

End Sub

